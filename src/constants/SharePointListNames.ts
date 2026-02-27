// ============================================================================
// DWx Policy Manager - SharePoint List Names
// Centralized configuration for all SharePoint list names
// ============================================================================

/**
 * Policy Management Lists
 * Core lists for policy lifecycle management
 */
export const PolicyLists = {
  /** Main policies list */
  POLICIES: 'PM_Policies',
  /** Policy version history */
  POLICY_VERSIONS: 'PM_PolicyVersions',
  /** User acknowledgements */
  POLICY_ACKNOWLEDGEMENTS: 'PM_PolicyAcknowledgements',
  /** Policy exemptions */
  POLICY_EXEMPTIONS: 'PM_PolicyExemptions',
  /** Distribution campaigns */
  POLICY_DISTRIBUTIONS: 'PM_PolicyDistributions',
  /** Policy templates */
  POLICY_TEMPLATES: 'PM_PolicyTemplates',
  /** User feedback on policies */
  POLICY_FEEDBACK: 'PM_PolicyFeedback',
  /** Audit trail */
  POLICY_AUDIT_LOG: 'PM_PolicyAuditLog',
  /** Policy analytics/metrics */
  POLICY_ANALYTICS: 'PM_PolicyAnalytics',
  /** Policy documents/attachments */
  POLICY_DOCUMENTS: 'PM_PolicyDocuments',
  /** Metadata profiles for policies */
  POLICY_METADATA_PROFILES: 'PM_PolicyMetadataProfiles',
  /** Policy reviewer assignments */
  POLICY_REVIEWERS: 'PM_PolicyReviewers',
  /** Read receipt tracking */
  POLICY_READ_RECEIPTS: 'PM_PolicyReadReceipts',
  /** Policy categorization */
  POLICY_CATEGORIES: 'PM_PolicyCategories',
  /** Policy requests from users */
  POLICY_REQUESTS: 'PM_PolicyRequests',
  /** Policy sub-categories for folder navigation */
  POLICY_SUB_CATEGORIES: 'PM_PolicySubCategories',
} as const;

/**
 * Quiz Lists
 * Lists for policy knowledge verification
 */
export const QuizLists = {
  /** Quiz definitions */
  POLICY_QUIZZES: 'PM_PolicyQuizzes',
  /** Quiz questions */
  POLICY_QUIZ_QUESTIONS: 'PM_PolicyQuizQuestions',
  /** Quiz attempt results */
  POLICY_QUIZ_RESULTS: 'PM_PolicyQuizResults',
  /** Quiz attempts (extended) */
  QUIZ_ATTEMPTS: 'PM_QuizAttempts',
  /** Question banks for reusable questions */
  QUESTION_BANKS: 'PM_QuestionBanks',
  /** Quiz sections for organization */
  QUIZ_SECTIONS: 'PM_QuizSections',
  /** Grading rubrics for manual grading */
  GRADING_RUBRICS: 'PM_GradingRubrics',
  /** Quiz certificates */
  QUIZ_CERTIFICATES: 'PM_QuizCertificates',
  /** Certificate templates */
  CERTIFICATE_TEMPLATES: 'PM_CertificateTemplates',
} as const;

/**
 * Social/Engagement Lists
 * Lists for policy engagement features
 */
export const SocialLists = {
  /** Policy ratings */
  POLICY_RATINGS: 'PM_PolicyRatings',
  /** Policy comments */
  POLICY_COMMENTS: 'PM_PolicyComments',
  /** Comment likes */
  POLICY_COMMENT_LIKES: 'PM_PolicyCommentLikes',
  /** Policy shares */
  POLICY_SHARES: 'PM_PolicyShares',
  /** Policy followers */
  POLICY_FOLLOWERS: 'PM_PolicyFollowers',
} as const;

/**
 * Policy Pack Lists
 * Lists for policy bundling and assignments
 */
export const PolicyPackLists = {
  /** Policy pack definitions */
  POLICY_PACKS: 'PM_PolicyPacks',
  /** Pack assignments to users */
  POLICY_PACK_ASSIGNMENTS: 'PM_PolicyPackAssignments',
} as const;

/**
 * Policy Workflow Lists
 * Lists for policy approval workflows
 */
export const PolicyWorkflowLists = {
  /** Workflow template definitions */
  WORKFLOW_TEMPLATES: 'PM_WorkflowTemplates',
  /** Active workflow instances */
  WORKFLOW_INSTANCES: 'PM_WorkflowInstances',
  /** Approval decision records */
  APPROVAL_DECISIONS: 'PM_ApprovalDecisions',
  /** Approval delegations */
  DELEGATIONS: 'PM_Delegations',
  /** Escalation rule definitions */
  ESCALATION_RULES: 'PM_EscalationRules',
  /** Workflow execution history */
  WORKFLOW_HISTORY: 'PM_WorkflowHistory',
} as const;

/**
 * Policy Retention Lists
 * Lists for retention and compliance management
 */
export const RetentionLists = {
  /** Retention policy definitions */
  RETENTION_POLICIES: 'PM_RetentionPolicies',
  /** Legal hold records */
  LEGAL_HOLDS: 'PM_LegalHolds',
  /** Archived policy records */
  RETENTION_ARCHIVE: 'PM_RetentionArchive',
} as const;

/**
 * Notification Lists
 * Lists for policy notifications and reminders
 */
export const NotificationLists = {
  /** Policy-specific notifications */
  POLICY_NOTIFICATIONS: 'PM_PolicyNotifications',
  /** Reminder schedules */
  REMINDER_SCHEDULE: 'PM_ReminderSchedule',
  /** In-app notifications */
  NOTIFICATIONS: 'PM_Notifications',
} as const;

/**
 * Policy Analytics Extended Lists
 * Additional lists for analytics and reporting
 */
export const AnalyticsLists = {
  /** User activity tracking */
  USER_ACTIVITY_LOG: 'PM_UserActivityLog',
  /** Compliance violation records */
  COMPLIANCE_VIOLATIONS: 'PM_ComplianceViolations',
  /** Department-level analytics */
  DEPARTMENT_ANALYTICS: 'PM_DepartmentAnalytics',
  /** Scheduled report definitions */
  SCHEDULED_REPORTS: 'PM_ScheduledReports',
  /** Report execution history */
  REPORT_EXECUTIONS: 'PM_ReportExecutions',
  /** Audit report records */
  AUDIT_REPORTS: 'PM_AuditReports',
  /** Audit trail entries */
  AUDIT_TRAIL: 'PM_AuditTrail',
} as const;

/**
 * Document Comparison Lists
 * Lists for policy document versioning and comparison
 */
export const ComparisonLists = {
  /** Policy version records */
  VERSIONS: 'PM_Versions',
  /** Comparison history */
  COMPARISON_HISTORY: 'PM_ComparisonHistory',
} as const;

/**
 * Template Library Lists
 * Lists for template management
 */
export const TemplateLibraryLists = {
  /** Template definitions */
  TEMPLATES: 'PM_Templates',
  /** Template usage tracking */
  TEMPLATE_USAGE: 'PM_TemplateUsage',
  /** User preferences for templates */
  USER_PREFERENCES: 'PM_UserPreferences',
  /** Corporate document templates */
  CORPORATE_TEMPLATES: 'PM_CorporateTemplates',
} as const;

/**
 * Admin Configuration Lists
 * Lists for admin panel settings and configuration
 */
export const AdminConfigLists = {
  /** Naming convention rules */
  NAMING_RULES: 'PM_NamingRules',
  /** SLA target configurations */
  SLA_CONFIGS: 'PM_SLAConfigs',
  /** Data lifecycle/retention policies */
  DATA_LIFECYCLE_POLICIES: 'PM_DataLifecyclePolicies',
  /** Email notification templates */
  EMAIL_TEMPLATES: 'PM_EmailTemplates',
} as const;

/**
 * User Management Lists
 * Lists for user directory, sync, and audience targeting
 */
export const UserManagementLists = {
  /** Employee directory (synced from Entra ID) */
  EMPLOYEES: 'PM_Employees',
  /** Entra ID sync operation logs */
  SYNC_LOG: 'PM_Sync_Log',
  /** Custom audience definitions for targeting */
  AUDIENCES: 'PM_Audiences',
} as const;

/**
 * System Lists
 * Lists for system operations
 */
export const SystemLists = {
  /** Notification queue */
  NOTIFICATION_QUEUE: 'PM_NotificationQueue',
  /** Audit archive */
  AUDIT_ARCHIVE: 'PM_PolicyAuditArchive',
  /** File conversion queue */
  FILE_CONVERSION_QUEUE: 'PM_FileConversionQueue',
  /** Policy source documents */
  POLICY_SOURCE_DOCUMENTS: 'PM_PolicySourceDocuments',
} as const;

/**
 * All Policy Manager Lists
 * Combined export of all PM_ lists
 */
export const PM_LISTS = {
  ...PolicyLists,
  ...QuizLists,
  ...SocialLists,
  ...PolicyPackLists,
  ...PolicyWorkflowLists,
  ...RetentionLists,
  ...NotificationLists,
  ...AnalyticsLists,
  ...ComparisonLists,
  ...TemplateLibraryLists,
  ...AdminConfigLists,
  ...UserManagementLists,
  ...SystemLists,
} as const;

/**
 * Legacy JML List Name Mapping
 * Maps old JML_ list names to new PM_ names
 * Use this for migration/reference only
 */
export const LegacyListMapping: Record<string, string> = {
  // Policy Core Lists
  'JML_Policies': PM_LISTS.POLICIES,
  'JML_Policy_Policies': PM_LISTS.POLICIES,
  'JML_PolicyVersions': PM_LISTS.POLICY_VERSIONS,
  'JML_Policy_Versions': PM_LISTS.VERSIONS,
  'JML_PolicyAcknowledgements': PM_LISTS.POLICY_ACKNOWLEDGEMENTS,
  'JML_Policy_Acknowledgements': PM_LISTS.POLICY_ACKNOWLEDGEMENTS,
  'JML_PolicyExemptions': PM_LISTS.POLICY_EXEMPTIONS,
  'JML_PolicyDistributions': PM_LISTS.POLICY_DISTRIBUTIONS,
  'JML_PolicyTemplates': PM_LISTS.POLICY_TEMPLATES,
  'JML_Policy_Templates': PM_LISTS.TEMPLATES,
  'JML_PolicyFeedback': PM_LISTS.POLICY_FEEDBACK,
  'JML_PolicyAuditLog': PM_LISTS.POLICY_AUDIT_LOG,
  'JML_PolicyAnalytics': PM_LISTS.POLICY_ANALYTICS,
  'JML_PolicyDocuments': PM_LISTS.POLICY_DOCUMENTS,
  // Quiz Lists
  'JML_PolicyQuizzes': PM_LISTS.POLICY_QUIZZES,
  'JML_Policy_Quizzes': PM_LISTS.POLICY_QUIZZES,
  'JML_PolicyQuizQuestions': PM_LISTS.POLICY_QUIZ_QUESTIONS,
  'JML_PolicyQuizResults': PM_LISTS.POLICY_QUIZ_RESULTS,
  'JML_Policy_QuizResults': PM_LISTS.POLICY_QUIZ_RESULTS,
  // Social Lists
  'JML_PolicyRatings': PM_LISTS.POLICY_RATINGS,
  'JML_PolicyComments': PM_LISTS.POLICY_COMMENTS,
  'JML_PolicyCommentLikes': PM_LISTS.POLICY_COMMENT_LIKES,
  'JML_PolicyShares': PM_LISTS.POLICY_SHARES,
  'JML_PolicyFollowers': PM_LISTS.POLICY_FOLLOWERS,
  // Policy Pack Lists
  'JML_PolicyPacks': PM_LISTS.POLICY_PACKS,
  'JML_PolicyPackAssignments': PM_LISTS.POLICY_PACK_ASSIGNMENTS,
  // Workflow Lists
  'JML_Policy_WorkflowTemplates': PM_LISTS.WORKFLOW_TEMPLATES,
  'JML_Policy_WorkflowInstances': PM_LISTS.WORKFLOW_INSTANCES,
  'JML_Policy_ApprovalDecisions': PM_LISTS.APPROVAL_DECISIONS,
  'JML_Policy_Delegations': PM_LISTS.DELEGATIONS,
  'JML_PolicyDelegations': PM_LISTS.DELEGATIONS,
  'JML_Policy_EscalationRules': PM_LISTS.ESCALATION_RULES,
  'JML_Policy_WorkflowHistory': PM_LISTS.WORKFLOW_HISTORY,
  // Retention Lists
  'JML_RetentionPolicies': PM_LISTS.RETENTION_POLICIES,
  'JML_LegalHolds': PM_LISTS.LEGAL_HOLDS,
  'JML_RetentionArchive': PM_LISTS.RETENTION_ARCHIVE,
  // Notification Lists
  'JML_PolicyNotifications': PM_LISTS.POLICY_NOTIFICATIONS,
  'JML_PolicyReminderSchedule': PM_LISTS.REMINDER_SCHEDULE,
  'JML_Notifications': PM_LISTS.NOTIFICATIONS,
  'JML_NotificationQueue': PM_LISTS.NOTIFICATION_QUEUE,
  // Analytics Lists
  'JML_UserActivityLog': PM_LISTS.USER_ACTIVITY_LOG,
  'JML_ComplianceViolations': PM_LISTS.COMPLIANCE_VIOLATIONS,
  'JML_DepartmentAnalytics': PM_LISTS.DEPARTMENT_ANALYTICS,
  'JML_ScheduledReports': PM_LISTS.SCHEDULED_REPORTS,
  'JML_ReportExecutions': PM_LISTS.REPORT_EXECUTIONS,
  'JML_AuditReports': PM_LISTS.AUDIT_REPORTS,
  'JML_AuditTrail': PM_LISTS.AUDIT_TRAIL,
  // Comparison Lists
  'JML_Policy_ComparisonHistory': PM_LISTS.COMPARISON_HISTORY,
  // Template Lists
  'JML_Policy_TemplateUsage': PM_LISTS.TEMPLATE_USAGE,
  'JML_Policy_UserPreferences': PM_LISTS.USER_PREFERENCES,
  // System Lists
  'JML_FileConversionQueue': PM_LISTS.FILE_CONVERSION_QUEUE,
  'JML_PolicySourceDocuments': PM_LISTS.POLICY_SOURCE_DOCUMENTS,
  // Additional Policy Lists
  'JML_PolicyMetadataProfiles': PM_LISTS.POLICY_METADATA_PROFILES,
  'JML_PolicyReviewers': PM_LISTS.POLICY_REVIEWERS,
  'JML_PolicyReadReceipts': PM_LISTS.POLICY_READ_RECEIPTS,
  'JML_PolicyCategories': PM_LISTS.POLICY_CATEGORIES,
};

// Type exports for TypeScript
export type PolicyListName = typeof PolicyLists[keyof typeof PolicyLists];
export type QuizListName = typeof QuizLists[keyof typeof QuizLists];
export type SocialListName = typeof SocialLists[keyof typeof SocialLists];
export type PolicyPackListName = typeof PolicyPackLists[keyof typeof PolicyPackLists];
export type PolicyWorkflowListName = typeof PolicyWorkflowLists[keyof typeof PolicyWorkflowLists];
export type RetentionListName = typeof RetentionLists[keyof typeof RetentionLists];
export type NotificationListName = typeof NotificationLists[keyof typeof NotificationLists];
export type AnalyticsListName = typeof AnalyticsLists[keyof typeof AnalyticsLists];
export type ComparisonListName = typeof ComparisonLists[keyof typeof ComparisonLists];
export type TemplateLibraryListName = typeof TemplateLibraryLists[keyof typeof TemplateLibraryLists];
export type AdminConfigListName = typeof AdminConfigLists[keyof typeof AdminConfigLists];
export type UserManagementListName = typeof UserManagementLists[keyof typeof UserManagementLists];
export type SystemListName = typeof SystemLists[keyof typeof SystemLists];
export type PMListName = typeof PM_LISTS[keyof typeof PM_LISTS];

/**
 * JML Integration Lists (Optional)
 * These lists may be used when integrating with JML system.
 * They remain with JML_ prefix as they belong to JML, not Policy Manager.
 */
export const JMLIntegrationLists = {
  /** JML Processes - for task assignment integration */
  PROCESSES: 'JML_Processes',
  /** JML Task Assignments */
  TASK_ASSIGNMENTS: 'JML_TaskAssignments',
  /** JML Employees */
  EMPLOYEES: 'JML_Employees',
} as const;
