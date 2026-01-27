// External Sharing Hub Models
// Interfaces for cross-tenant collaboration and B2B guest management

import { IBaseListItem, IUser } from './ICommon';

/**
 * Trust Level for partner organizations
 */
export enum TrustLevel {
  Full = 'Full',
  Limited = 'Limited',
  Custom = 'Custom'
}

/**
 * Trust Relationship Status
 */
export enum TrustStatus {
  Active = 'Active',
  Pending = 'Pending',
  Suspended = 'Suspended',
  Revoked = 'Revoked'
}

/**
 * Guest User Status
 */
export enum GuestStatus {
  Active = 'Active',
  Suspended = 'Suspended',
  Removed = 'Removed'
}

/**
 * Guest Invitation Status
 */
export enum InvitationStatus {
  PendingAcceptance = 'PendingAcceptance',
  Accepted = 'Accepted',
  Expired = 'Expired',
  Failed = 'Failed'
}

/**
 * Risk Level for security assessment
 */
export enum RiskLevel {
  Low = 'Low',
  Medium = 'Medium',
  High = 'High'
}

/**
 * Shared Resource Type
 */
export enum ResourceType {
  Document = 'Document',
  Folder = 'Folder',
  Site = 'Site',
  List = 'List',
  Library = 'Library'
}

/**
 * Sharing Level (Permission Level)
 */
export enum SharingLevel {
  View = 'View',
  Edit = 'Edit',
  FullControl = 'Full Control'
}

/**
 * Shared Resource Status
 */
export enum SharedResourceStatus {
  Active = 'Active',
  Revoked = 'Revoked',
  Expired = 'Expired'
}

/**
 * Data Classification Level for External Sharing
 * Note: This enum is specific to External Sharing with HighlyConfidential option
 */
export enum DataClassification {
  Public = 'Public',
  Internal = 'Internal',
  Confidential = 'Confidential',
  HighlyConfidential = 'Highly Confidential'
}

/**
 * Acknowledgment Status
 */
export enum AcknowledgmentStatus {
  Pending = 'Pending',
  Acknowledged = 'Acknowledged',
  Declined = 'Declined'
}

/**
 * Related JML Module
 */
export enum RelatedModule {
  Contract = 'Contract',
  Procurement = 'Procurement',
  Signing = 'Signing',
  Policy = 'Policy',
  Other = 'Other'
}

/**
 * Audit Action Type
 */
export enum AuditActionType {
  TrustEstablished = 'TrustEstablished',
  TrustRevoked = 'TrustRevoked',
  TrustModified = 'TrustModified',
  GuestInvited = 'GuestInvited',
  GuestRemoved = 'GuestRemoved',
  GuestSuspended = 'GuestSuspended',
  GuestReactivated = 'GuestReactivated',
  ResourceShared = 'ResourceShared',
  ResourceUnshared = 'ResourceUnshared',
  PermissionChanged = 'PermissionChanged',
  AccessReviewCompleted = 'AccessReviewCompleted',
  PolicyViolation = 'PolicyViolation'
}

/**
 * Audit Result
 */
export enum AuditResult {
  Success = 'Success',
  Failure = 'Failure',
  Warning = 'Warning'
}

/**
 * Access Review Status
 */
export enum AccessReviewStatus {
  Pending = 'Pending',
  InProgress = 'InProgress',
  Completed = 'Completed',
  Expired = 'Expired'
}

// Alias for backward compatibility
export type ReviewStatus = AccessReviewStatus;
export const ReviewStatus = AccessReviewStatus;

/**
 * Access Review Type
 */
export enum AccessReviewType {
  Guest = 'Guest',
  Resource = 'Resource',
  Organization = 'Organization'
}

// Alias for backward compatibility
export type ReviewType = AccessReviewType;
export const ReviewType = AccessReviewType;

/**
 * Review Decision
 */
export enum ReviewDecision {
  Approve = 'Approve',
  Revoke = 'Revoke',
  Modify = 'Modify'
}

/**
 * Action Taken after review
 */
export enum ActionTaken {
  None = 'None',
  AccessRevoked = 'AccessRevoked',
  PermissionsModified = 'PermissionsModified',
  Approved = 'Approved'
}

/**
 * Access Level for guest users
 */
export enum AccessLevel {
  Guest = 'Guest',
  Member = 'Member',
  Limited = 'Limited'
}

/**
 * Policy Type
 */
export enum SharingPolicyType {
  Organization = 'Organization',
  Resource = 'Resource',
  User = 'User'
}

// ============================================
// SharePoint List Interfaces
// ============================================

/**
 * Trusted Organization - JML_TrustedOrganizations list
 */
export interface ITrustedOrganization extends IBaseListItem {
  TenantId: string;
  TenantDomain: string;
  TrustLevel: TrustLevel;
  Status: TrustStatus;
  TrustMFAClaims: boolean;
  TrustDeviceClaims: boolean;
  TrustHybridJoinedDevices: boolean;
  AllowedDomains?: string; // JSON array
  AllowedUserGroups?: string; // JSON array
  DefaultGuestExpiration: number; // Days
  MaxSharingLevel: SharingLevel;
  InboundAccessEnabled: boolean;
  OutboundAccessEnabled: boolean;
  B2BDirectConnectEnabled: boolean;
  ContactName?: string;
  ContactEmail?: string;
  Notes?: string;
  EstablishedDate?: Date;
  LastVerifiedDate?: Date;
  VerifiedById?: number;
  VerifiedBy?: IUser;
}

/**
 * External Guest User - JML_ExternalGuestUsers list
 */
export interface IExternalGuestUser extends IBaseListItem {
  Email: string;
  UserPrincipalName?: string;
  AzureADObjectId?: string;
  SourceOrganizationId?: number;
  SourceOrganization?: ITrustedOrganization;
  InvitedById?: number;
  InvitedBy?: IUser;
  InvitationDate?: Date;
  InvitationStatus: InvitationStatus;
  FirstAccessDate?: Date;
  LastAccessDate?: Date;
  AccessExpirationDate?: Date;
  IsExpired: boolean;
  AccessLevel: 'Guest' | 'Member' | 'Limited';
  AssignedSites?: string; // JSON array
  AssignedGroups?: string; // JSON array
  TotalResourcesAccessed: number;
  MFARegistered: boolean;
  DeviceCompliant: boolean;
  RiskLevel: RiskLevel;
  Status: GuestStatus;
  SuspensionReason?: string;
  Notes?: string;
}

/**
 * Shared Resource - JML_ExternalSharedResources list
 */
export interface ISharedResource extends IBaseListItem {
  ResourceType: ResourceType;
  ResourceUrl: string;
  ResourceId?: string;
  SharedWithOrganizationId?: number;
  SharedWithOrganization?: ITrustedOrganization;
  SharedWithUsers?: string; // JSON array of external users
  SharingLevel: SharingLevel;
  SharedDate: Date;
  SharedById?: number;
  SharedBy?: IUser;
  ExpirationDate?: Date;
  IsExpired: boolean;
  DataClassification: DataClassification;
  RequiresAcknowledgment: boolean;
  AcknowledgmentStatus: AcknowledgmentStatus;
  RelatedModule: RelatedModule;
  RelatedItemId?: string;
  AccessCount: number;
  LastAccessedDate?: Date;
  LastAccessedBy?: string;
  Status: SharedResourceStatus;
}

/**
 * Audit Log Entry for External Sharing - JML_ExternalSharingAuditLog list
 * Note: Named differently to avoid conflict with IAuditLogEntry from ICommon
 */
export interface IExternalSharingAuditLog extends IBaseListItem {
  ActionType: AuditActionType;
  PerformedById?: number;
  PerformedBy?: IUser;
  PerformedDate: Date;
  TargetOrganizationId?: number;
  TargetOrganization?: ITrustedOrganization;
  TargetResourceId?: number;
  TargetResource?: ISharedResource;
  TargetUser?: string; // External user email
  PreviousValue?: string; // JSON
  NewValue?: string; // JSON
  IPAddress?: string;
  UserAgent?: string;
  Result: AuditResult;
  CorrelationId?: string;
  RiskScore?: number; // 0-100
}

// Type alias for backward compatibility
export type IAuditLogEntryExternal = IExternalSharingAuditLog;

/**
 * Sharing Policy - JML_ExternalSharingPolicies list
 */
export interface ISharingPolicy extends IBaseListItem {
  PolicyType: SharingPolicyType;
  IsDefault: boolean;
  IsActive: boolean;
  Scope?: string; // JSON defining where policy applies
  MaxGuestExpiration: number; // Days
  RequireMFA: boolean;
  RequireDeviceCompliance: boolean;
  AllowedFileTypes?: string; // JSON array
  BlockedFileTypes?: string; // JSON array
  MaxFileSizeMB?: number;
  RequireDataClassification: boolean;
  MinDataClassification?: DataClassification;
  AllowAnonymousLinks: boolean;
  AllowGuestDownload: boolean;
  AllowGuestEdit: boolean;
  RequireAccessReview: boolean;
  AccessReviewFrequencyDays?: number;
  EnforceWatermark: boolean;
  NotifyOnShare: boolean;
  NotifyRecipients?: string; // JSON array of emails
  ApprovedById?: number;
  ApprovedBy?: IUser;
  ApprovedDate?: Date;
}

/**
 * Access Review - JML_ExternalAccessReviews list
 */
export interface IAccessReview extends IBaseListItem {
  ReviewType: AccessReviewType;
  TargetId: string;
  TargetType: string;
  ReviewerEmail: string;
  ReviewerId?: number;
  Reviewer?: IUser;
  Status: AccessReviewStatus;
  DueDate: Date | string;
  CompletedDate?: Date | string;
  Decision?: ReviewDecision;
  Justification?: string;
  ActionTaken?: ActionTaken;
  ActionDate?: Date | string;
  NextReviewDate?: Date | string;
  AutoRevokeOnExpiry: boolean;
}

// ============================================
// Service Request/Response Interfaces
// ============================================

/**
 * Trust Configuration Request
 */
export interface ITrustConfig {
  organizationName: string;
  tenantId: string;
  tenantDomain: string;
  trustLevel: TrustLevel;
  trustMFAClaims: boolean;
  trustDeviceClaims: boolean;
  trustHybridJoinedDevices: boolean;
  allowedDomains?: string[];
  allowedUserGroups?: string[];
  defaultGuestExpiration: number;
  maxSharingLevel: SharingLevel;
  inboundAccessEnabled: boolean;
  outboundAccessEnabled: boolean;
  b2bDirectConnectEnabled: boolean;
  contactName?: string;
  contactEmail?: string;
  notes?: string;
}

/**
 * Trust Configuration (wizard version with all fields)
 */
export interface ITrustConfiguration extends ITrustConfig {
  // All fields from ITrustConfig plus these are the same
  // This type alias exists for semantic clarity in wizard components
}

/**
 * Trust Health Status
 */
export interface ITrustHealthStatus {
  organizationId: number;
  organizationName: string;
  tenantDomain: string;
  isHealthy: boolean;
  lastVerifiedDate?: Date;
  policyInSync: boolean;
  activeGuestCount: number;
  activeResourceCount: number;
  pendingReviews: number;
  riskLevel: RiskLevel;
  issues: string[];
}

/**
 * Guest Filter for queries
 */
export interface IGuestFilter {
  organizationId?: number;
  status?: GuestStatus;
  invitationStatus?: InvitationStatus;
  riskLevel?: RiskLevel;
  isExpired?: boolean;
  searchText?: string;
}

/**
 * Guest Invitation Request
 */
export interface IInvitation {
  email: string;
  displayName: string;
  sourceOrganizationId: number;
  sendInvitationMessage: boolean;
  invitationMessage?: string;
  accessExpirationDays?: number;
  accessLevel?: AccessLevel;
  expirationDate?: Date;
}

/**
 * Guest Invitation Result
 */
export interface IInvitationResult {
  success: boolean;
  guestId?: number;
  azureAdObjectId?: string;
  invitationUrl?: string;
  error?: string;
}

/**
 * Bulk Invite Result
 */
export interface IBulkInviteResult {
  total: number;
  succeeded: number;
  failed: number;
  results: IInvitationResult[];
}

/**
 * Guest Access Details
 */
export interface IGuestAccessDetails {
  guest: IExternalGuestUser;
  resources: ISharedResource[];
  recentActivity: IExternalSharingAuditLog[];
  pendingReviews: IAccessReview[];
}

/**
 * Resource Filter for queries
 */
export interface IResourceFilter {
  organizationId?: number;
  resourceType?: ResourceType;
  status?: SharedResourceStatus;
  relatedModule?: RelatedModule;
  isExpired?: boolean;
  sharedById?: number;
  searchText?: string;
}

/**
 * Share Resource Request
 */
export interface IShareRequest {
  title: string;
  resourceType: ResourceType;
  resourceUrl: string;
  resourceId?: string;
  sharedWithOrganizationId?: number;
  sharedWithUsers?: string[];
  sharingLevel: SharingLevel;
  expirationDays?: number;
  expirationDate?: Date;
  dataClassification: DataClassification;
  requiresAcknowledgment: boolean;
  relatedModule: RelatedModule | 'Contract' | 'Procurement' | 'Signing' | 'Policy' | 'Other';
  relatedItemId?: string;
  message?: string;
}

/**
 * Bulk Share Result
 */
export interface IBulkShareResult {
  total: number;
  succeeded: number;
  failed: number;
  results: {
    resourceUrl: string;
    success: boolean;
    resourceId?: number;
    error?: string;
  }[];
}

/**
 * Access Log Entry for resources
 */
export interface IAccessLogEntry {
  date: Date;
  userEmail: string;
  userName?: string;
  action: 'Viewed' | 'Downloaded' | 'Edited' | 'Shared';
  ipAddress?: string;
  deviceInfo?: string;
}

/**
 * Audit Filter for queries
 */
export interface IAuditFilter {
  startDate?: Date;
  endDate?: Date;
  actionTypes?: AuditActionType[];
  organizationId?: number;
  performedById?: number;
  result?: AuditResult;
  targetUser?: string;
  minRiskScore?: number;
}

/**
 * Compliance Report Options
 */
export interface IReportOptions {
  reportType: 'Summary' | 'Detailed' | 'AuditTrail' | 'RiskAssessment';
  startDate: Date;
  endDate: Date;
  organizationIds?: number[];
  includeCharts: boolean;
  format: 'PDF' | 'Excel' | 'JSON';
}

/**
 * Compliance Report
 */
export interface IComplianceReport {
  reportId: string;
  reportType: string;
  generatedDate: Date;
  periodStart: Date;
  periodEnd: Date;
  summary: {
    totalTrustedOrgs: number;
    totalGuests: number;
    totalSharedResources: number;
    totalAuditEvents: number;
    complianceScore: number;
    riskScore: number;
  };
  findings: {
    critical: string[];
    warnings: string[];
    recommendations: string[];
  };
  data?: unknown;
}

/**
 * Risk Context for scoring
 */
export interface IRiskContext {
  userId?: string;
  organizationId?: number;
  resourceId?: number;
  action?: AuditActionType;
  ipAddress?: string;
  userAgent?: string;
  isNewLocation?: boolean;
  isOffHours?: boolean;
  recentFailedAttempts?: number;
}

/**
 * Security Alert
 */
export interface ISecurityAlert {
  id: string;
  alertType: 'SuspiciousActivity' | 'PolicyViolation' | 'ExpiredAccess' | 'HighRisk' | 'AnomalousAccess';
  severity: 'Low' | 'Medium' | 'High' | 'Critical';
  title: string;
  description: string;
  timestamp: Date;
  organizationId?: number;
  userId?: string;
  resourceId?: number;
  isResolved: boolean;
  resolvedById?: number;
  resolvedDate?: Date;
  resolution?: string;
}

/**
 * Access Review Request
 */
export interface IAccessReviewRequest {
  reviewType: AccessReviewType;
  targetId: string;
  targetType: string;
  reviewerEmail: string;
  dueDays: number;
  autoRevokeOnExpiry: boolean;
}

/**
 * Review Decision Request
 */
export interface IReviewDecision {
  decision: ReviewDecision;
  justification: string;
  newExpirationDate?: Date;
  newPermissionLevel?: SharingLevel;
}

/**
 * Review Policy Configuration
 */
export interface IReviewPolicy {
  reviewType: AccessReviewType;
  frequencyDays: number;
  reviewerSelection: 'Manager' | 'ResourceOwner' | 'Specific';
  specificReviewers?: string[];
  autoRevokeOnExpiry: boolean;
  reminderDays: number[];
}

// ============================================
// Cross-Tenant Access Policy Interfaces (Graph API)
// ============================================

/**
 * Cross-Tenant Access Policy from Graph API
 */
export interface ICrossTenantPolicy {
  id: string;
  displayName?: string;
  description?: string;
  default: {
    b2bCollaborationInbound?: ICrossTenantAccessSettings;
    b2bCollaborationOutbound?: ICrossTenantAccessSettings;
    b2bDirectConnectInbound?: ICrossTenantAccessSettings;
    b2bDirectConnectOutbound?: ICrossTenantAccessSettings;
    inboundTrust?: IInboundTrustSettings;
  };
  partners: IPartnerConfiguration[];
}

/**
 * Partner Configuration
 */
export interface IPartnerConfiguration {
  tenantId: string;
  displayName?: string;
  b2bCollaborationInbound?: ICrossTenantAccessSettings;
  b2bCollaborationOutbound?: ICrossTenantAccessSettings;
  b2bDirectConnectInbound?: ICrossTenantAccessSettings;
  b2bDirectConnectOutbound?: ICrossTenantAccessSettings;
  inboundTrust?: IInboundTrustSettings;
  isServiceProvider?: boolean;
  isInMultiTenantOrganization?: boolean;
}

/**
 * Cross-Tenant Access Settings
 */
export interface ICrossTenantAccessSettings {
  usersAndGroups?: {
    accessType: 'allowed' | 'blocked';
    targets: {
      target: string; // 'AllUsers' or specific group/user ID
      targetType: 'user' | 'group';
    }[];
  };
  applications?: {
    accessType: 'allowed' | 'blocked';
    targets: {
      target: string; // 'AllApplications' or specific app ID
      targetType: 'application';
    }[];
  };
}

/**
 * Inbound Trust Settings
 */
export interface IInboundTrustSettings {
  isMfaAccepted: boolean;
  isCompliantDeviceAccepted: boolean;
  isHybridAzureADJoinedDeviceAccepted: boolean;
}

// ============================================
// Dashboard & KPI Interfaces
// ============================================

/**
 * External Sharing KPIs
 */
export interface IExternalSharingKPIs {
  activeTrustedOrganizations: number;
  pendingTrustedOrganizations: number;
  activeGuestUsers: number;
  expiringGuestUsers: number;
  activeSharedResources: number;
  expiringResources: number;
  pendingAccessReviews: number;
  overdueAccessReviews: number;
  securityAlerts: number;
  complianceScore: number;
  riskScore: number;
}

/**
 * Activity Feed Item
 */
export interface IActivityFeedItem {
  id: string;
  type: AuditActionType;
  title: string;
  description: string;
  timestamp: Date;
  performedBy: string;
  organizationName?: string;
  resourceName?: string;
  isHighRisk: boolean;
}

/**
 * External User in selection context
 */
export interface IExternalRecipient {
  email: string;
  displayName: string;
  organizationId: number;
  organizationName: string;
  isNew: boolean; // True if not yet a guest
}

/**
 * Validation Result
 */
export interface IValidationResult {
  isValid: boolean;
  errors: string[];
  warnings: string[];
}
