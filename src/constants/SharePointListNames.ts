// ============================================================================
// DWx Policy Manager — SharePoint List Names
// Centralized configuration for all SharePoint list names
//
// ORGANISATION:
//   Section 1: ACTIVE — Lists currently used by the application
//   Section 2: PLANNED — Lists for confirmed future features (V2+)
//   Section 3: Combined PM_LISTS export (Active + Planned)
//
// Last audited: 29 Mar 2026 (Session 20 — Production Hardening)
// ============================================================================

// ============================================================================
// SECTION 1: ACTIVE LISTS — Currently used in production code
// ============================================================================

/**
 * Core Policy Lists
 * The backbone of policy lifecycle management
 */
export const PolicyLists = {
  /** Main policies list — all policy records */
  POLICIES: 'PM_Policies',
  /** Policy version history — snapshots on publish/revise */
  POLICY_VERSIONS: 'PM_PolicyVersions',
  /** User acknowledgements — tracks who read/acknowledged each policy */
  POLICY_ACKNOWLEDGEMENTS: 'PM_PolicyAcknowledgements',
  /** Distribution campaigns — bulk policy distribution tracking */
  POLICY_DISTRIBUTIONS: 'PM_PolicyDistributions',
  /** Policy templates — reusable policy structures */
  POLICY_TEMPLATES: 'PM_PolicyTemplates',
  /** User feedback — support requests from Help Centre */
  POLICY_FEEDBACK: 'PM_PolicyFeedback',
  /** Audit trail — all policy actions logged for compliance */
  POLICY_AUDIT_LOG: 'PM_PolicyAuditLog',
  /** Metadata profiles — Fast Track template presets */
  POLICY_METADATA_PROFILES: 'PM_PolicyMetadataProfiles',
  /** Reviewer assignments — who reviews each policy */
  POLICY_REVIEWERS: 'PM_PolicyReviewers',
  /** Policy categories — category definitions with sort order */
  POLICY_CATEGORIES: 'PM_PolicyCategories',
  /** Policy requests — user-submitted policy creation requests */
  POLICY_REQUESTS: 'PM_PolicyRequests',
  /** Sub-categories — nested category tree under parent categories */
  POLICY_SUB_CATEGORIES: 'PM_PolicySubCategories',
  /** Source documents — document library with per-policy folders */
  POLICY_SOURCE_DOCUMENTS: 'PM_PolicySourceDocuments',
  /** Policy exemptions */
  POLICY_EXEMPTIONS: 'PM_PolicyExemptions',
  /** Policy documents/attachments (legacy — use POLICY_SOURCE_DOCUMENTS) */
  POLICY_DOCUMENTS: 'PM_PolicyDocuments',
  /** Policy analytics/metrics */
  POLICY_ANALYTICS: 'PM_PolicyAnalytics',
  /** Distribution queue — server-side bulk processing */
  DISTRIBUTION_QUEUE: 'PM_DistributionQueue',
  /** Read receipt tracking */
  POLICY_READ_RECEIPTS: 'PM_PolicyReadReceipts',
  /** Policy-specific notifications (legacy — use NOTIFICATIONS) */
  POLICY_NOTIFICATIONS: 'PM_PolicyNotifications',
} as const;

/**
 * Quiz Lists
 * Policy knowledge verification system
 */
export const QuizLists = {
  /** Quiz definitions — settings, passing score, time limit */
  POLICY_QUIZZES: 'PM_PolicyQuizzes',
  /** Quiz questions — individual questions with type-specific data */
  POLICY_QUIZ_QUESTIONS: 'PM_PolicyQuizQuestions',
  /** Quiz results — aggregated scores per user per quiz */
  POLICY_QUIZ_RESULTS: 'PM_PolicyQuizResults',
  /** Quiz attempts — individual attempt records with timestamps */
  QUIZ_ATTEMPTS: 'PM_QuizAttempts',
  /** Question banks — reusable question pools across quizzes */
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
 * Policy Pack Lists
 * Bundle policies together for streamlined distribution
 */
export const PolicyPackLists = {
  /** Policy pack definitions — name, description, policies, settings */
  POLICY_PACKS: 'PM_PolicyPacks',
  /** Pack assignments — who was assigned each pack */
  POLICY_PACK_ASSIGNMENTS: 'PM_PolicyPackAssignments',
} as const;

/**
 * Approval Lists
 * Policy approval workflow — currently used via direct SP writes
 */
export const ApprovalLists = {
  /** Individual approval records */
  APPROVALS: 'PM_Approvals',
  /** Approval action audit trail — who approved/rejected, when, why */
  APPROVAL_HISTORY: 'PM_ApprovalHistory',
  /** Delegation assignments — delegate approvals to another user */
  APPROVAL_DELEGATIONS: 'PM_ApprovalDelegations',
} as const;

/**
 * Notification Lists
 * In-app notifications and email delivery queue
 */
export const NotificationLists = {
  /** In-app notifications — displayed in notification bell */
  NOTIFICATIONS: 'PM_Notifications',
  /** Email/notification delivery queue — polled by Logic App */
  NOTIFICATION_QUEUE: 'PM_NotificationQueue',
  /** Reminder schedule — automated reminders for ack/review/expiry */
  REMINDER_SCHEDULE: 'PM_ReminderSchedule',
  /** Policy-specific notifications (legacy alias for NOTIFICATIONS) */
  POLICY_NOTIFICATIONS: 'PM_PolicyNotifications',
} as const;

/**
 * Admin & Configuration Lists
 * System configuration and user management
 */
export const AdminLists = {
  /** Key-value configuration store — all admin settings */
  CONFIGURATION: 'PM_Configuration',
  /** User profiles — synced from Entra ID, used for audience targeting */
  USER_PROFILES: 'PM_UserProfiles',
  /** Event Viewer diagnostic log — persisted Error/Critical events + manual saves */
  EVENT_LOG: 'PM_EventLog',
} as const;

// ============================================================================
// SECTION 2: PLANNED LISTS — Confirmed for future releases (V2+)
// These constants are defined so provisioning scripts and services can
// reference them, but the UI features are not yet wired.
// ============================================================================

/**
 * Social & Engagement Lists — PLANNED (V2)
 * Policy ratings, comments, sharing, and following.
 * Backend: PolicySocialService exists with full CRUD methods.
 * UI: Needs rating widget on PolicyDetails, comments section, share button.
 */
export const SocialLists = {
  /** 5-star policy ratings */
  POLICY_RATINGS: 'PM_PolicyRatings',
  /** Discussion comments on policies */
  POLICY_COMMENTS: 'PM_PolicyComments',
  /** Likes on comments */
  POLICY_COMMENT_LIKES: 'PM_PolicyCommentLikes',
  /** Share tracking — who shared which policy */
  POLICY_SHARES: 'PM_PolicyShares',
  /** Policy followers — users who want update notifications */
  POLICY_FOLLOWERS: 'PM_PolicyFollowers',
} as const;

/**
 * Workflow Engine Lists — PLANNED (V2)
 * Multi-level approval chains, auto-escalation, workflow templates.
 * Backend: ApprovalService has full engine code (unused by UI).
 * Current: UI writes directly to PM_Approvals + PM_ApprovalHistory.
 * V2: Wire ApprovalService for multi-level chains + escalation.
 */
export const WorkflowLists = {
  /** Reusable workflow templates (Fast Track, Standard, Regulatory) */
  WORKFLOW_TEMPLATES: 'PM_WorkflowTemplates',
  /** Active workflow instances — running approval chains */
  WORKFLOW_INSTANCES: 'PM_WorkflowInstances',
  /** Approval chain instances — level progression tracking */
  APPROVAL_CHAINS: 'PM_ApprovalChains',
  /** Reusable approval workflow templates */
  APPROVAL_TEMPLATES: 'PM_ApprovalTemplates',
  /** Escalation rule definitions — auto-reassign after timeout */
  ESCALATION_RULES: 'PM_EscalationRules',
  /** Workflow execution history — full audit of chain progression */
  WORKFLOW_HISTORY: 'PM_WorkflowHistory',
} as const;

/**
 * Retention & Compliance Lists — PLANNED (V2)
 * Retention policies, legal holds, SLA breach logging.
 * Current: Data Lifecycle in Admin uses PM_Configuration key-value.
 * V2: Dedicated lists for per-policy retention + legal hold tracking.
 */
export const RetentionLists = {
  /** Retention policy definitions per document type */
  RETENTION_POLICIES: 'PM_RetentionPolicies',
  /** Legal hold records — prevents deletion during litigation */
  LEGAL_HOLDS: 'PM_LegalHolds',
  /** Archived policy records — moved from PM_Policies after retention */
  RETENTION_ARCHIVE: 'PM_RetentionArchive',
  /** SLA breach records — persisted for compliance audit trail */
  SLA_BREACHES: 'PM_SLABreaches',
} as const;

/**
 * Extended Analytics Lists — PLANNED (V2)
 * Department-level analytics, scheduled reports, activity tracking.
 * Current: Analytics computed client-side from core lists.
 * V2: Server-side aggregation for large-scale deployments.
 */
export const AnalyticsLists = {
  /** User activity tracking — page views, downloads, time spent */
  USER_ACTIVITY_LOG: 'PM_UserActivityLog',
  /** Department-level compliance metrics */
  DEPARTMENT_ANALYTICS: 'PM_DepartmentAnalytics',
  /** Report definitions — saved report configurations */
  REPORT_DEFINITIONS: 'PM_ReportDefinitions',
  /** Scheduled report configurations — auto-run + email */
  SCHEDULED_REPORTS: 'PM_ScheduledReports',
  /** Report execution history — when reports ran, who received them */
  REPORT_EXECUTIONS: 'PM_ReportExecutions',
} as const;

/**
 * User Management Extended Lists — PLANNED (V2)
 * Entra ID sync logging and audience rule definitions.
 * Current: AudienceRuleService evaluates from PM_UserProfiles directly.
 * V2: Dedicated audience definitions + sync audit trail.
 */
export const UserManagementLists = {
  /** Entra ID sync operation logs */
  SYNC_LOG: 'PM_Sync_Log',
  /** Custom audience definitions for targeting */
  AUDIENCES: 'PM_Audiences',
} as const;

// ============================================================================
// SECTION 3: COMBINED EXPORT
// All list constants available via PM_LISTS.CONSTANT_NAME
// ============================================================================

/**
 * Legacy group aliases — some services import these directly
 */
export const AdminConfigLists = {
  ...AdminLists,
  POLICY_CATEGORIES: PolicyLists.POLICY_CATEGORIES,
  NAMING_RULES: 'PM_NamingRules',
  SLA_CONFIGS: 'PM_SLAConfigs',
  DATA_LIFECYCLE_POLICIES: 'PM_DataLifecyclePolicies',
  EMAIL_TEMPLATES: 'PM_EmailTemplates',
} as const;
export const SystemLists = {
  ...NotificationLists,
  POLICY_SOURCE_DOCUMENTS: PolicyLists.POLICY_SOURCE_DOCUMENTS,
  POLICY_NOTIFICATIONS: PolicyLists.POLICY_NOTIFICATIONS,
  AUDIT_ARCHIVE: 'PM_PolicyAuditArchive',
  FILE_CONVERSION_QUEUE: 'PM_FileConversionQueue',
} as const;
export const PolicyWorkflowLists = {
  ...ApprovalLists,
  ...WorkflowLists,
  DELEGATIONS: ApprovalLists.APPROVAL_DELEGATIONS,
  APPROVAL_DECISIONS: 'PM_ApprovalDecisions',
} as const;

export const PM_LISTS = {
  // Active
  ...PolicyLists,
  ...QuizLists,
  ...PolicyPackLists,
  ...ApprovalLists,
  ...NotificationLists,
  ...AdminLists,
  // Planned (V2+)
  ...SocialLists,
  ...WorkflowLists,
  ...RetentionLists,
  ...AnalyticsLists,
  ...UserManagementLists,
} as const;

// ============================================================================
// TYPE EXPORTS
// ============================================================================

export type PMListName = typeof PM_LISTS[keyof typeof PM_LISTS];

// ============================================================================
// LEGACY JML MAPPING (migration reference only)
// ============================================================================

export const LegacyListMapping: Record<string, string> = {
  'JML_Policies': PM_LISTS.POLICIES,
  'JML_Policy_Policies': PM_LISTS.POLICIES,
  'JML_PolicyVersions': PM_LISTS.POLICY_VERSIONS,
  'JML_PolicyAcknowledgements': PM_LISTS.POLICY_ACKNOWLEDGEMENTS,
  'JML_PolicyDistributions': PM_LISTS.POLICY_DISTRIBUTIONS,
  'JML_PolicyTemplates': PM_LISTS.POLICY_TEMPLATES,
  'JML_PolicyFeedback': PM_LISTS.POLICY_FEEDBACK,
  'JML_PolicyAuditLog': PM_LISTS.POLICY_AUDIT_LOG,
  'JML_PolicyDocuments': PM_LISTS.POLICY_SOURCE_DOCUMENTS,
  'JML_PolicyQuizzes': PM_LISTS.POLICY_QUIZZES,
  'JML_PolicyQuizQuestions': PM_LISTS.POLICY_QUIZ_QUESTIONS,
  'JML_PolicyQuizResults': PM_LISTS.POLICY_QUIZ_RESULTS,
  'JML_PolicyRatings': PM_LISTS.POLICY_RATINGS,
  'JML_PolicyComments': PM_LISTS.POLICY_COMMENTS,
  'JML_PolicyPacks': PM_LISTS.POLICY_PACKS,
  'JML_PolicyPackAssignments': PM_LISTS.POLICY_PACK_ASSIGNMENTS,
  'JML_Notifications': PM_LISTS.NOTIFICATIONS,
  'JML_NotificationQueue': PM_LISTS.NOTIFICATION_QUEUE,
  'JML_PolicyMetadataProfiles': PM_LISTS.POLICY_METADATA_PROFILES,
  'JML_PolicyReviewers': PM_LISTS.POLICY_REVIEWERS,
  'JML_PolicyCategories': PM_LISTS.POLICY_CATEGORIES,
  'JML_PolicySourceDocuments': PM_LISTS.POLICY_SOURCE_DOCUMENTS,
};
