/**
 * User Sync Configuration Models
 *
 * Comprehensive configuration for Entra ID user synchronization
 * including notifications, mapping rules, and delta sync settings.
 */

/**
 * Complete sync configuration
 */
export interface IUserSyncConfig {
  /** General sync settings */
  general: IGeneralSyncSettings;
  /** Email notification settings */
  notifications: INotificationSettings;
  /** User mapping rules */
  mappingRules: IUserMappingRule[];
  /** Delta sync settings */
  deltaSync: IDeltaSyncSettings;
  /** Schedule settings */
  schedule: ISyncScheduleConfig;
}

/**
 * General sync settings
 */
export interface IGeneralSyncSettings {
  /** Include disabled Entra users */
  includeDisabledUsers: boolean;
  /** Update existing employees */
  updateExisting: boolean;
  /** Deactivate employees not in Entra */
  deactivateMissing: boolean;
  /** User types to sync */
  userTypes: ('Member' | 'Guest')[];
  /** Departments to include (empty = all) */
  departmentFilter: string[];
  /** Users to exclude by UPN */
  excludeUsers: string[];
  /** Batch size for processing */
  batchSize: number;
}

/**
 * Email notification settings
 */
export interface INotificationSettings {
  /** Enable email notifications */
  enabled: boolean;
  /** Send on successful sync */
  onSuccess: boolean;
  /** Send on sync with errors */
  onError: boolean;
  /** Send on sync failure */
  onFailure: boolean;
  /** Include detailed report */
  includeDetailedReport: boolean;
  /** Recipients for notifications */
  recipients: INotificationRecipient[];
  /** Email template settings */
  template: IEmailTemplate;
}

/**
 * Notification recipient
 */
export interface INotificationRecipient {
  /** Email address */
  email: string;
  /** Display name */
  name: string;
  /** Notification types to receive */
  notifyOn: ('success' | 'error' | 'failure')[];
}

/**
 * Email template configuration
 */
export interface IEmailTemplate {
  /** Subject line template */
  subject: string;
  /** Include stats summary */
  includeStats: boolean;
  /** Include error details */
  includeErrors: boolean;
  /** Include list of added users */
  includeAddedUsers: boolean;
  /** Max users to list in email */
  maxUsersToList: number;
  /** Custom footer text */
  footerText: string;
}

/**
 * User mapping rule for auto-assignment
 */
export interface IUserMappingRule {
  /** Unique rule ID */
  id: string;
  /** Rule name */
  name: string;
  /** Rule description */
  description?: string;
  /** Is rule active */
  isActive: boolean;
  /** Rule priority (lower = higher priority) */
  priority: number;
  /** Conditions to match */
  conditions: IUserMappingCondition[];
  /** Match type for conditions */
  conditionMatch: 'all' | 'any';
  /** Actions to take when matched */
  actions: IUserMappingAction[];
}

/**
 * Condition for user mapping
 */
export interface IUserMappingCondition {
  /** Field to check */
  field: UserMappingField;
  /** Operator for comparison */
  operator: MappingOperator;
  /** Value to compare */
  value: string;
  /** Case sensitive comparison */
  caseSensitive: boolean;
}

/**
 * Fields available for mapping conditions
 */
export type UserMappingField =
  | 'department'
  | 'jobTitle'
  | 'officeLocation'
  | 'companyName'
  | 'userPrincipalName'
  | 'mail'
  | 'displayName'
  | 'employeeType'
  | 'employeeId';

/**
 * Operators for mapping conditions
 */
export type MappingOperator =
  | 'equals'
  | 'notEquals'
  | 'contains'
  | 'notContains'
  | 'startsWith'
  | 'endsWith'
  | 'matches'  // regex
  | 'isEmpty'
  | 'isNotEmpty';

/**
 * Action to take when rule matches
 */
export interface IUserMappingAction {
  /** Action type */
  type: MappingActionType;
  /** Target field or value */
  target: string;
  /** Value to set */
  value: string;
}

/**
 * Types of mapping actions
 */
export type MappingActionType =
  | 'assignRole'       // Add to JML role group
  | 'setField'         // Set a field value
  | 'addToGroup'       // Add to SharePoint group
  | 'setEmploymentType'
  | 'setStatus'
  | 'setCostCenter'
  | 'skip';            // Skip this user

/**
 * Delta sync settings
 */
export interface IDeltaSyncSettings {
  /** Enable delta sync */
  enabled: boolean;
  /** Delta token for Graph API */
  deltaToken?: string;
  /** Last delta sync timestamp */
  lastDeltaSync?: Date;
  /** Fallback to full sync if delta fails */
  fallbackToFull: boolean;
  /** Force full sync every N delta syncs */
  forceFullSyncEvery: number;
  /** Delta sync counter */
  deltaSyncCount: number;
}

/**
 * Sync schedule configuration
 */
export interface ISyncScheduleConfig {
  /** Schedule enabled */
  enabled: boolean;
  /** Frequency */
  frequency: 'hourly' | 'daily' | 'weekly' | 'monthly';
  /** Time of day (HH:MM) */
  timeOfDay: string;
  /** Day of week (0-6, 0=Sunday) */
  dayOfWeek?: number;
  /** Day of month (1-31) */
  dayOfMonth?: number;
  /** Timezone */
  timezone: string;
  /** Last run */
  lastRun?: Date;
  /** Next scheduled run */
  nextRun?: Date;
}

/**
 * Sync analytics data
 */
export interface ISyncAnalytics {
  /** Summary statistics */
  summary: ISyncSummaryStats;
  /** Trend data */
  trends: ISyncTrendData;
  /** Department breakdown */
  byDepartment: IDepartmentSyncStats[];
  /** Recent sync operations */
  recentSyncs: ISyncOperationSummary[];
  /** Error analysis */
  errorAnalysis: IErrorAnalysis;
}

/**
 * Summary statistics
 */
export interface ISyncSummaryStats {
  /** Total syncs all time */
  totalSyncs: number;
  /** Syncs this month */
  syncsThisMonth: number;
  /** Syncs this week */
  syncsThisWeek: number;
  /** Total users synced */
  totalUsersSynced: number;
  /** Success rate percentage */
  successRate: number;
  /** Average sync duration (seconds) */
  avgDuration: number;
  /** Last successful sync */
  lastSuccessfulSync?: Date;
  /** Users added this month */
  usersAddedThisMonth: number;
  /** Users updated this month */
  usersUpdatedThisMonth: number;
}

/**
 * Trend data for charts
 */
export interface ISyncTrendData {
  /** Daily sync counts for last 30 days */
  dailySyncs: { date: string; count: number; success: number; failed: number }[];
  /** Weekly user changes */
  weeklyUserChanges: { week: string; added: number; updated: number; deactivated: number }[];
  /** Monthly totals */
  monthlyTotals: { month: string; syncs: number; users: number; errors: number }[];
}

/**
 * Department sync statistics
 */
export interface IDepartmentSyncStats {
  /** Department name */
  department: string;
  /** Total employees */
  totalEmployees: number;
  /** Active employees */
  activeEmployees: number;
  /** Last synced */
  lastSynced?: Date;
  /** Sync coverage percentage */
  syncCoverage: number;
}

/**
 * Sync operation summary
 */
export interface ISyncOperationSummary {
  /** Sync ID */
  syncId: string;
  /** Timestamp */
  timestamp: Date;
  /** Type (Full/Delta/Filtered) */
  type: string;
  /** Status */
  status: string;
  /** Duration in seconds */
  duration: number;
  /** Users processed */
  usersProcessed: number;
  /** Added */
  added: number;
  /** Updated */
  updated: number;
  /** Errors */
  errors: number;
}

/**
 * Error analysis
 */
export interface IErrorAnalysis {
  /** Total errors this month */
  totalErrors: number;
  /** Error by type */
  byType: { type: string; count: number; lastOccurred: Date }[];
  /** Most common errors */
  topErrors: { message: string; count: number; affectedUsers: string[] }[];
  /** Error trend */
  trend: 'increasing' | 'decreasing' | 'stable';
}

/**
 * Default sync configuration
 */
export const DEFAULT_SYNC_CONFIG: IUserSyncConfig = {
  general: {
    includeDisabledUsers: true,
    updateExisting: true,
    deactivateMissing: false,
    userTypes: ['Member'],
    departmentFilter: [],
    excludeUsers: [],
    batchSize: 50
  },
  notifications: {
    enabled: false,
    onSuccess: true,
    onError: true,
    onFailure: true,
    includeDetailedReport: false,
    recipients: [],
    template: {
      subject: 'JML User Sync {status} - {date}',
      includeStats: true,
      includeErrors: true,
      includeAddedUsers: false,
      maxUsersToList: 20,
      footerText: 'This is an automated message from JML User Sync Service.'
    }
  },
  mappingRules: [],
  deltaSync: {
    enabled: false,
    fallbackToFull: true,
    forceFullSyncEvery: 7,
    deltaSyncCount: 0
  },
  schedule: {
    enabled: false,
    frequency: 'daily',
    timeOfDay: '06:00',
    timezone: 'UTC'
  }
};

/**
 * Example mapping rules
 */
export const EXAMPLE_MAPPING_RULES: IUserMappingRule[] = [
  {
    id: 'hr-department-rule',
    name: 'HR Department Auto-Role',
    description: 'Automatically assign HR Admin role to HR department members',
    isActive: true,
    priority: 1,
    conditions: [
      {
        field: 'department',
        operator: 'equals',
        value: 'Human Resources',
        caseSensitive: false
      }
    ],
    conditionMatch: 'all',
    actions: [
      {
        type: 'assignRole',
        target: 'HRAdmin',
        value: 'JML HR Admins'
      }
    ]
  },
  {
    id: 'it-department-rule',
    name: 'IT Department Auto-Role',
    description: 'Automatically assign IT Admin role to IT department members',
    isActive: true,
    priority: 2,
    conditions: [
      {
        field: 'department',
        operator: 'contains',
        value: 'IT',
        caseSensitive: false
      }
    ],
    conditionMatch: 'all',
    actions: [
      {
        type: 'assignRole',
        target: 'ITAdmin',
        value: 'JML IT Admins'
      }
    ]
  },
  {
    id: 'manager-title-rule',
    name: 'Manager Title Auto-Role',
    description: 'Assign Line Manager role to users with Manager in title',
    isActive: true,
    priority: 3,
    conditions: [
      {
        field: 'jobTitle',
        operator: 'contains',
        value: 'Manager',
        caseSensitive: false
      }
    ],
    conditionMatch: 'all',
    actions: [
      {
        type: 'assignRole',
        target: 'LineManager',
        value: 'JML Line Managers'
      }
    ]
  },
  {
    id: 'contractor-type-rule',
    name: 'Contractor Employment Type',
    description: 'Set employment type for contractors',
    isActive: true,
    priority: 4,
    conditions: [
      {
        field: 'employeeType',
        operator: 'equals',
        value: 'Contractor',
        caseSensitive: false
      }
    ],
    conditionMatch: 'all',
    actions: [
      {
        type: 'setEmploymentType',
        target: 'EmploymentType',
        value: 'Contractor'
      }
    ]
  }
];
