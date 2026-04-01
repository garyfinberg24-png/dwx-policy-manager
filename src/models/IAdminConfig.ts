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
  category?: string;
  subject: string;
  body: string;
  recipients: string;
  isActive: boolean;
  isDefault?: boolean;
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
// CUSTOM THEME
// ============================================================================

export interface ICustomTheme {
  // Branding
  logoUrl: string;
  logoText: string;
  tagline: string;
  footerText: string;
  faviconUrl: string;

  // Colors
  primaryColor: string;
  primaryDark: string;
  accentColor: string;
  successColor: string;
  warningColor: string;
  dangerColor: string;

  // Header
  headerStyle: 'solid' | 'gradient';
  headerGradientStart: string;
  headerGradientEnd: string;

  // Surfaces
  sidebarBackground: string;
  contentBackground: string;
  cardBackground: string;

  // Typography
  fontFamily: string;

  // Borders
  cardBorderRadius: number;
  controlBorderRadius: number;

  // Preset
  preset: string;
}

export const DEFAULT_THEME: ICustomTheme = {
  logoUrl: '',
  logoText: 'Policy Manager',
  tagline: 'POLICY GOVERNANCE & COMPLIANCE',
  footerText: 'DWx Digital Workplace. All rights reserved.',
  faviconUrl: '',
  primaryColor: '#0d9488',
  primaryDark: '#0f766e',
  accentColor: '#0284c7',
  successColor: '#059669',
  warningColor: '#d97706',
  dangerColor: '#dc2626',
  headerStyle: 'gradient',
  headerGradientStart: '#0d9488',
  headerGradientEnd: '#0f766e',
  sidebarBackground: '#f1f5f9',
  contentBackground: '#ffffff',
  cardBackground: '#ffffff',
  fontFamily: 'Segoe UI',
  cardBorderRadius: 4,
  controlBorderRadius: 4,
  preset: 'forest-teal'
};

export const PRESET_THEMES: Record<string, Partial<ICustomTheme>> = {
  'forest-teal': {
    preset: 'forest-teal',
    primaryColor: '#0d9488', primaryDark: '#0f766e',
    headerGradientStart: '#0d9488', headerGradientEnd: '#0f766e',
    accentColor: '#0284c7', successColor: '#059669', warningColor: '#d97706', dangerColor: '#dc2626'
  },
  'corporate-blue': {
    preset: 'corporate-blue',
    primaryColor: '#1e40af', primaryDark: '#1e3a8a',
    headerGradientStart: '#1e40af', headerGradientEnd: '#1e3a8a',
    accentColor: '#0369a1', successColor: '#059669', warningColor: '#d97706', dangerColor: '#dc2626'
  },
  'slate-professional': {
    preset: 'slate-professional',
    primaryColor: '#475569', primaryDark: '#334155',
    headerGradientStart: '#475569', headerGradientEnd: '#334155',
    accentColor: '#0284c7', successColor: '#059669', warningColor: '#d97706', dangerColor: '#dc2626'
  },
  'royal-purple': {
    preset: 'royal-purple',
    primaryColor: '#7c3aed', primaryDark: '#6d28d9',
    headerGradientStart: '#7c3aed', headerGradientEnd: '#6d28d9',
    accentColor: '#2563eb', successColor: '#059669', warningColor: '#d97706', dangerColor: '#dc2626'
  },
  'crimson-red': {
    preset: 'crimson-red',
    primaryColor: '#dc2626', primaryDark: '#b91c1c',
    headerGradientStart: '#dc2626', headerGradientEnd: '#b91c1c',
    accentColor: '#0284c7', successColor: '#059669', warningColor: '#d97706', dangerColor: '#7c3aed'
  },
  'forest-green': {
    preset: 'forest-green',
    primaryColor: '#15803d', primaryDark: '#166534',
    headerGradientStart: '#15803d', headerGradientEnd: '#166534',
    accentColor: '#0284c7', successColor: '#0d9488', warningColor: '#d97706', dangerColor: '#dc2626'
  },
  'midnight': {
    preset: 'midnight',
    primaryColor: '#1e293b', primaryDark: '#0f172a',
    headerGradientStart: '#1e293b', headerGradientEnd: '#0f172a',
    accentColor: '#3b82f6', successColor: '#10b981', warningColor: '#f59e0b', dangerColor: '#ef4444',
    sidebarBackground: '#1e293b', contentBackground: '#f8fafc'
  },
  'microsoft-fluent': {
    preset: 'microsoft-fluent',
    primaryColor: '#0078d4', primaryDark: '#106ebe',
    headerGradientStart: '#0078d4', headerGradientEnd: '#106ebe',
    accentColor: '#0078d4', successColor: '#107c10', warningColor: '#ffb900', dangerColor: '#d13438',
    sidebarBackground: '#f3f2f1', contentBackground: '#ffffff'
  }
};

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

  // Event Viewer Settings
  static readonly EVENTVIEWER_ENABLED = 'Admin.EventViewer.Enabled';
  static readonly EVENTVIEWER_APP_BUFFER_SIZE = 'Admin.EventViewer.AppBufferSize';
  static readonly EVENTVIEWER_CONSOLE_BUFFER_SIZE = 'Admin.EventViewer.ConsoleBufferSize';
  static readonly EVENTVIEWER_NETWORK_BUFFER_SIZE = 'Admin.EventViewer.NetworkBufferSize';
  static readonly EVENTVIEWER_AUTO_PERSIST_THRESHOLD = 'Admin.EventViewer.AutoPersistThreshold';
  static readonly EVENTVIEWER_AI_TRIAGE_ENABLED = 'Admin.EventViewer.AITriageEnabled';
  static readonly EVENTVIEWER_AI_FUNCTION_URL = 'Admin.EventViewer.AIFunctionUrl';
  static readonly EVENTVIEWER_RETENTION_DAYS = 'Admin.EventViewer.RetentionDays';
  static readonly EVENTVIEWER_HIDE_CDN_DEFAULT = 'Admin.EventViewer.HideCDNByDefault';

  // Performance Optimizer Settings (written by Event Viewer sliders)
  static readonly PERF_CACHE_TTL = 'Perf.CacheTTL';
  static readonly PERF_REQUEST_DEDUP = 'Perf.RequestDedup';
  static readonly PERF_LEAN_QUERIES = 'Perf.LeanQueries';
  static readonly PERF_DEFAULT_TOP_LIMIT = 'Perf.DefaultTopLimit';
  static readonly PERF_MAX_CONCURRENT = 'Perf.MaxConcurrent';
}
