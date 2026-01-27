// IT Provisioning Models
// Interfaces for Azure AD/Entra ID, M365 License, and Group Management

import { IUser } from './ICommon';

// ============================================================================
// Entra ID / Azure AD Types
// ============================================================================

export interface IEntraUserCreateRequest {
  displayName: string;
  givenName: string;
  surname: string;
  mailNickname: string;
  userPrincipalName: string;
  mail?: string;
  jobTitle?: string;
  department?: string;
  officeLocation?: string;
  companyName?: string;
  employeeId?: string;
  mobilePhone?: string;
  businessPhones?: string[];
  streetAddress?: string;
  city?: string;
  state?: string;
  country?: string;
  postalCode?: string;
  usageLocation: string; // Required for license assignment (e.g., "US", "GB")
  managerId?: string; // Entra ID of manager
  accountEnabled: boolean;
  passwordProfile: IPasswordProfile;
}

export interface IPasswordProfile {
  password: string;
  forceChangePasswordNextSignIn: boolean;
  forceChangePasswordNextSignInWithMfa?: boolean;
}

export interface IEntraUser {
  id: string;
  displayName: string;
  givenName?: string;
  surname?: string;
  userPrincipalName: string;
  mail?: string;
  jobTitle?: string;
  department?: string;
  officeLocation?: string;
  companyName?: string;
  employeeId?: string;
  mobilePhone?: string;
  accountEnabled: boolean;
  createdDateTime?: Date;
  signInSessionsValidFromDateTime?: Date;
  manager?: {
    id: string;
    displayName: string;
  };
}

export interface IEntraUserUpdateRequest {
  displayName?: string;
  givenName?: string;
  surname?: string;
  jobTitle?: string;
  department?: string;
  officeLocation?: string;
  mobilePhone?: string;
  streetAddress?: string;
  city?: string;
  state?: string;
  country?: string;
  postalCode?: string;
  managerId?: string;
  accountEnabled?: boolean;
}

// ============================================================================
// M365 License Types (Graph API)
// ============================================================================

export interface IGraphLicense {
  skuId: string;
  skuPartNumber: string;
  servicePlans: IServicePlan[];
  consumedUnits: number;
  prepaidUnits: {
    enabled: number;
    suspended: number;
    warning: number;
  };
}

export interface IServicePlan {
  servicePlanId: string;
  servicePlanName: string;
  provisioningStatus: 'Success' | 'PendingInput' | 'PendingActivation' | 'PendingProvisioning' | 'Disabled';
  appliesTo: 'User' | 'Company';
}

export interface ILicenseAssignmentRequest {
  userId: string;
  addLicenses: Array<{
    skuId: string;
    disabledPlans?: string[]; // Service plans to disable
  }>;
  removeLicenses: string[]; // SKU IDs to remove
}

export interface IUserLicenseDetail {
  userId: string;
  userPrincipalName: string;
  displayName: string;
  assignedLicenses: Array<{
    skuId: string;
    skuPartNumber: string;
    disabledPlans: string[];
  }>;
  licenseAssignmentStates: Array<{
    skuId: string;
    state: 'Active' | 'ActiveWithError' | 'Disabled' | 'Error';
    error?: string;
  }>;
}

// ============================================================================
// Group Management Types
// ============================================================================

export interface ISecurityGroup {
  id: string;
  displayName: string;
  description?: string;
  mail?: string;
  mailEnabled: boolean;
  securityEnabled: boolean;
  groupTypes: string[];
  membershipRule?: string;
  membershipRuleProcessingState?: 'On' | 'Paused';
  visibility?: 'Private' | 'Public' | 'HiddenMembership';
}

export interface IGroupMembershipChange {
  groupId: string;
  userId: string;
  action: 'add' | 'remove';
  groupDisplayName?: string;
  userDisplayName?: string;
}

// ============================================================================
// Teams Provisioning Types
// ============================================================================

export interface ITeamMembershipChange {
  teamId: string;
  userId: string;
  action: 'add' | 'remove';
  role?: 'member' | 'owner';
  teamDisplayName?: string;
  userDisplayName?: string;
}

export interface ITeamInfo {
  id: string;
  displayName: string;
  description?: string;
  visibility?: 'Private' | 'Public';
  webUrl?: string;
}

// ============================================================================
// Provisioning Workflow Types
// ============================================================================

export type ProvisioningActionType =
  | 'CreateUser'
  | 'DisableUser'
  | 'EnableUser'
  | 'UpdateUser'
  | 'DeleteUser'
  | 'AssignLicense'
  | 'RemoveLicense'
  | 'AddToGroup'
  | 'RemoveFromGroup'
  | 'AddToTeam'
  | 'RemoveFromTeam'
  | 'SendWelcomeEmail'
  | 'SetOutOfOffice'
  | 'RevokeSession';

export type ProvisioningStatus =
  | 'Pending'
  | 'InProgress'
  | 'Completed'
  | 'Failed'
  | 'PartiallyCompleted'
  | 'RolledBack';

export interface IProvisioningStep {
  id: string;
  name: string;
  actionType: ProvisioningActionType;
  status: ProvisioningStatus;
  order: number;
  targetResource: string; // User ID, Group ID, etc.
  targetResourceName?: string;
  requestPayload?: string; // JSON
  responsePayload?: string; // JSON
  errorMessage?: string;
  startedAt?: Date;
  completedAt?: Date;
  canRollback: boolean;
  rollbackCompleted?: boolean;
}

export interface IProvisioningRequest {
  id?: number;
  processId: number;
  employeeId?: string;
  employeeName: string;
  employeeEmail: string;
  processType: 'Joiner' | 'Mover' | 'Leaver';
  department: string;
  jobTitle: string;
  manager?: IUser;
  startDate?: Date;
  endDate?: Date; // For leavers
  status: ProvisioningStatus;
  steps: IProvisioningStep[];
  createdAt: Date;
  completedAt?: Date;
  createdById: number;
  createdByName?: string;
  notes?: string;
}

export interface IProvisioningResult {
  success: boolean;
  requestId: number;
  status: ProvisioningStatus;
  completedSteps: number;
  totalSteps: number;
  failedStep?: string;
  errorMessage?: string;
  userCreated?: IEntraUser;
  licensesAssigned?: string[];
  groupsAdded?: string[];
  teamsAdded?: string[];
}

// ============================================================================
// Provisioning Configuration Types
// ============================================================================

export interface IDepartmentProvisioningConfig {
  department: string;
  defaultLicenses: string[]; // SKU IDs
  securityGroups: string[]; // Group IDs
  teams: string[]; // Team IDs
  sharePointSites?: string[]; // Site URLs
}

export interface IRoleProvisioningConfig {
  role: string;
  additionalLicenses?: string[];
  additionalGroups?: string[];
  additionalTeams?: string[];
}

export interface IProvisioningConfig {
  tenantId: string;
  defaultUsageLocation: string;
  passwordLength: number;
  forcePasswordChange: boolean;
  sendWelcomeEmail: boolean;
  welcomeEmailTemplate?: string;
  departmentConfigs: IDepartmentProvisioningConfig[];
  roleConfigs: IRoleProvisioningConfig[];
  leaverGracePeriodDays: number; // Days before license removal
  autoDisableOnLeave: boolean;
}

// ============================================================================
// Audit Log Types
// ============================================================================

export interface IProvisioningAuditLog {
  Id?: number;
  ProcessId: number;
  RequestId: number;
  EmployeeId?: string;
  EmployeeName: string;
  ActionType: ProvisioningActionType;
  ActionStatus: 'Success' | 'Failed' | 'RolledBack';
  TargetResource: string;
  TargetResourceName?: string;
  RequestPayload?: string;
  ResponsePayload?: string;
  ErrorDetails?: string;
  ExecutedById: number;
  ExecutedByName?: string;
  ExecutedAt: Date;
  RollbackAt?: Date;
  IPAddress?: string;
  UserAgent?: string;
}

// ============================================================================
// Common License SKU IDs (for reference)
// ============================================================================

export const CommonLicenseSkus = {
  // Microsoft 365
  M365_E3: 'ENTERPRISEPACK',
  M365_E5: 'ENTERPRISEPREMIUM',
  M365_F1: 'DESKLESSPACK',
  M365_F3: 'SPE_F1',
  M365_BUSINESS_BASIC: 'O365_BUSINESS_ESSENTIALS',
  M365_BUSINESS_STANDARD: 'O365_BUSINESS_PREMIUM',
  M365_BUSINESS_PREMIUM: 'SPB',

  // Office 365
  O365_E1: 'STANDARDPACK',
  O365_E3: 'ENTERPRISEPACK',
  O365_E5: 'ENTERPRISEPREMIUM',

  // Exchange Online
  EXCHANGE_ONLINE_PLAN_1: 'EXCHANGESTANDARD',
  EXCHANGE_ONLINE_PLAN_2: 'EXCHANGEENTERPRISE',

  // SharePoint
  SHAREPOINT_ONLINE_PLAN_1: 'SHAREPOINTSTANDARD',
  SHAREPOINT_ONLINE_PLAN_2: 'SHAREPOINTENTERPRISE',

  // Teams
  TEAMS_ESSENTIALS: 'TEAMS_ESSENTIALS',

  // Power Platform
  POWER_BI_PRO: 'POWER_BI_PRO',
  POWER_APPS_PER_USER: 'POWERAPPS_PER_USER',
  POWER_AUTOMATE_PER_USER: 'POWER_AUTOMATE_PER_USER',

  // Visio & Project
  VISIO_PLAN_2: 'VISIO_CLIENT_SUBSCRIPTION',
  PROJECT_PLAN_3: 'PROJECTPROFESSIONAL',

  // Enterprise Mobility
  EMS_E3: 'EMS',
  EMS_E5: 'EMSPREMIUM',

  // Azure AD Premium
  AAD_PREMIUM_P1: 'AAD_PREMIUM',
  AAD_PREMIUM_P2: 'AAD_PREMIUM_P2'
} as const;

export type CommonLicenseSku = typeof CommonLicenseSkus[keyof typeof CommonLicenseSkus];
