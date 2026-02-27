// ============================================================================
// DWx Policy Manager - Admin Configuration Interfaces
// Shared model types for admin panel sections
// ============================================================================

// ============================================================================
// NAMING RULES
// ============================================================================

export interface INamingRuleSegment {
  id: string;
  type: 'prefix' | 'counter' | 'date' | 'category' | 'separator' | 'freetext';
  value: string;
  format?: string;
}

export interface INamingRule {
  Id: number;
  Title: string;
  Pattern: string;
  Segments: INamingRuleSegment[];
  AppliesTo: string;
  IsActive: boolean;
  Example: string;
}

// ============================================================================
// SLA CONFIGS
// ============================================================================

export interface ISLAConfig {
  Id: number;
  Title: string;
  ProcessType: string;
  TargetDays: number;
  WarningThresholdDays: number;
  IsActive: boolean;
  Description: string;
}

// ============================================================================
// DATA LIFECYCLE / RETENTION
// ============================================================================

export interface IDataLifecyclePolicy {
  Id: number;
  Title: string;
  EntityType: string;
  RetentionPeriodDays: number;
  AutoDeleteEnabled: boolean;
  ArchiveBeforeDelete: boolean;
  IsActive: boolean;
  Description: string;
}

// ============================================================================
// EMAIL TEMPLATES
// Component uses lowercase field names for state; service maps to/from SP columns
// ============================================================================

export interface IEmailTemplate {
  id: number;
  name: string;
  event: string;
  subject: string;
  body: string;
  recipients: string;
  isActive: boolean;
  lastModified: string;
  mergeTags: string[];
}

// ============================================================================
// GENERAL SETTINGS
// ============================================================================

export interface IGeneralSettings {
  showFeaturedPolicy: boolean;
  showRecentlyViewed: boolean;
  showQuickStats: boolean;
  defaultViewMode: 'table' | 'card';
  policiesPerPage: number;
  enableSocialFeatures: boolean;
  enablePolicyRatings: boolean;
  enablePolicyComments: boolean;
  maintenanceMode: boolean;
  maintenanceMessage: string;
  aiFunctionUrl: string;
}

// ============================================================================
// NAVIGATION TOGGLE
// ============================================================================

export interface INavToggleItem {
  key: string;
  label: string;
  icon: string;
  description: string;
  isVisible: boolean;
}

// ============================================================================
// POLICY CATEGORIES
// ============================================================================

export interface IPolicyCategory {
  Id: number;
  Title: string;
  CategoryName: string;
  IconName: string;
  Color: string;
  Description: string;
  SortOrder: number;
  IsActive: boolean;
  IsDefault: boolean;
}

// ============================================================================
// SUB-CATEGORIES
// ============================================================================

export interface IPolicySubCategory {
  Id: number;
  Title: string;
  SubCategoryName: string;
  ParentCategoryId: number;
  ParentCategoryName: string;
  IconName: string;
  Description: string;
  SortOrder: number;
  IsActive: boolean;
}

// ============================================================================
// METADATA PROFILES
// ============================================================================

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
  IsDefault?: boolean;
  IsActive?: boolean;
  Description?: string;
}

// ============================================================================
// CONFIG KEY CONSTANTS for PM_Configuration key-value pairs
// ============================================================================

export class AdminConfigKeys {
  // General Settings
  static readonly GENERAL_SHOW_FEATURED = 'Admin.General.ShowFeaturedPolicy';
  static readonly GENERAL_SHOW_RECENTLY_VIEWED = 'Admin.General.ShowRecentlyViewed';
  static readonly GENERAL_SHOW_QUICK_STATS = 'Admin.General.ShowQuickStats';
  static readonly GENERAL_DEFAULT_VIEW = 'Admin.General.DefaultViewMode';
  static readonly GENERAL_POLICIES_PER_PAGE = 'Admin.General.PoliciesPerPage';
  static readonly GENERAL_SOCIAL_FEATURES = 'Admin.General.EnableSocialFeatures';
  static readonly GENERAL_POLICY_RATINGS = 'Admin.General.EnablePolicyRatings';
  static readonly GENERAL_POLICY_COMMENTS = 'Admin.General.EnablePolicyComments';
  static readonly GENERAL_MAINTENANCE_MODE = 'Admin.General.MaintenanceMode';
  static readonly GENERAL_MAINTENANCE_MESSAGE = 'Admin.General.MaintenanceMessage';

  // Approval Workflow Settings
  static readonly APPROVAL_REQUIRE_NEW = 'Admin.Approval.RequireForNew';
  static readonly APPROVAL_REQUIRE_UPDATE = 'Admin.Approval.RequireForUpdate';
  static readonly APPROVAL_ALLOW_SELF = 'Admin.Approval.AllowSelfApproval';

  // Compliance Settings
  static readonly COMPLIANCE_REQUIRE_ACK = 'Admin.Compliance.RequireAcknowledgement';
  static readonly COMPLIANCE_DEFAULT_DEADLINE = 'Admin.Compliance.DefaultDeadlineDays';
  static readonly COMPLIANCE_SEND_REMINDERS = 'Admin.Compliance.SendReminders';
  static readonly COMPLIANCE_REVIEW_FREQUENCY = 'Admin.Compliance.DefaultReviewFrequency';
  static readonly COMPLIANCE_REVIEW_REMINDERS = 'Admin.Compliance.SendReviewReminders';

  // Notification Settings
  static readonly NOTIFY_NEW_POLICIES = 'Admin.Notifications.NewPolicies';
  static readonly NOTIFY_POLICY_UPDATES = 'Admin.Notifications.PolicyUpdates';
  static readonly NOTIFY_DAILY_DIGEST = 'Admin.Notifications.DailyDigest';

  // Security Settings
  static readonly SECURITY_SESSION_TIMEOUT = 'Admin.Security.SessionTimeout';
  static readonly SECURITY_PASSWORD_POLICY = 'Admin.Security.PasswordPolicy';
  static readonly SECURITY_MFA_REQUIRED = 'Admin.Security.MFARequired';
  static readonly SECURITY_IP_LOGGING = 'Admin.Security.IPLogging';
  static readonly SECURITY_SENSITIVE_ACCESS_ALERTS = 'Admin.Security.SensitiveAccessAlerts';
  static readonly SECURITY_BULK_EXPORT_NOTIFY = 'Admin.Security.BulkExportNotifications';
  static readonly SECURITY_FAILED_LOGIN_LOCKOUT = 'Admin.Security.FailedLoginLockout';
}
