/**
 * Entra ID (Azure AD) User Sync Models
 * Used for synchronizing users from Entra ID to PM_Employees list
 */

/**
 * Entra ID user profile from Microsoft Graph
 */
export interface IEntraUser {
  /** Unique identifier in Entra ID */
  id: string;
  /** User principal name (usually email) */
  userPrincipalName: string;
  /** Display name */
  displayName: string;
  /** First name */
  givenName?: string;
  /** Last name */
  surname?: string;
  /** Primary email address */
  mail?: string;
  /** Job title */
  jobTitle?: string;
  /** Department */
  department?: string;
  /** Office location */
  officeLocation?: string;
  /** Business phone numbers */
  businessPhones?: string[];
  /** Mobile phone */
  mobilePhone?: string;
  /** Manager's Entra ID */
  managerId?: string;
  /** Employee ID (from HR system) */
  employeeId?: string;
  /** Employee type */
  employeeType?: string;
  /** Account enabled status */
  accountEnabled?: boolean;
  /** User type (Member, Guest) */
  userType?: string;
  /** Created date */
  createdDateTime?: string;
  /** Company name */
  companyName?: string;
  /** Cost center (extension attribute) */
  costCenter?: string;
  /** Profile photo URL */
  photoUrl?: string;
}

/**
 * JML Employee record (SharePoint list item)
 */
export interface IJMLEmployee {
  /** SharePoint item ID */
  Id?: number;
  /** Full name (Title field) */
  Title: string;
  /** First name */
  FirstName?: string;
  /** Last name */
  LastName?: string;
  /** Email address */
  Email: string;
  /** Employee number */
  EmployeeNumber?: string;
  /** Job title */
  JobTitle?: string;
  /** Department */
  Department?: string;
  /** Office location */
  Location?: string;
  /** Office phone */
  OfficePhone?: string;
  /** Mobile phone */
  MobilePhone?: string;
  /** Manager (Person field ID) */
  ManagerId?: number;
  /** Start date */
  StartDate?: Date;
  /** End date */
  EndDate?: Date;
  /** Employee status */
  Status: EmployeeStatus;
  /** Employment type */
  EmploymentType?: EmploymentType;
  /** Cost center */
  CostCenter?: string;
  /** Entra ID Object ID (for sync matching) */
  EntraObjectId?: string;
  /** Profile photo URL */
  ProfilePhoto?: string;
  /** Notes */
  Notes?: string;
  /** Last sync timestamp */
  LastSyncedAt?: Date;
}

/**
 * Employee status options
 */
export type EmployeeStatus =
  | 'Active'
  | 'Inactive'
  | 'PreHire'
  | 'OnLeave'
  | 'Terminated'
  | 'Retired';

/**
 * Employment type options
 */
export type EmploymentType =
  | 'Full-Time'
  | 'Part-Time'
  | 'Contractor'
  | 'Intern'
  | 'Temporary';

/**
 * Sync operation result for a single user
 */
export interface ISyncResult {
  /** User email or identifier */
  userIdentifier: string;
  /** Display name */
  displayName: string;
  /** Operation performed */
  operation: SyncOperation;
  /** Whether operation succeeded */
  success: boolean;
  /** Error message if failed */
  error?: string;
  /** SharePoint item ID (for adds/updates) */
  itemId?: number;
}

/**
 * Sync operation types
 */
export type SyncOperation =
  | 'Added'
  | 'Updated'
  | 'Deactivated'
  | 'Skipped'
  | 'Error';

/**
 * Overall sync job summary
 */
export interface ISyncSummary {
  /** Sync job ID */
  syncId: string;
  /** Start time */
  startedAt: Date;
  /** End time */
  completedAt?: Date;
  /** Sync status */
  status: SyncStatus;
  /** Total users processed */
  totalProcessed: number;
  /** Users added */
  added: number;
  /** Users updated */
  updated: number;
  /** Users deactivated */
  deactivated: number;
  /** Users skipped */
  skipped: number;
  /** Errors encountered */
  errors: number;
  /** Individual results */
  results: ISyncResult[];
  /** Error details */
  errorDetails?: string[];
}

/**
 * Sync job status
 */
export type SyncStatus =
  | 'Running'
  | 'Completed'
  | 'CompletedWithErrors'
  | 'Failed';

/**
 * Sync configuration options
 */
export interface ISyncConfig {
  /** Include disabled Entra users (mark as Inactive) */
  includeDisabledUsers: boolean;
  /** Update existing employees with latest Entra data */
  updateExisting: boolean;
  /** Deactivate employees not found in Entra */
  deactivateMissing: boolean;
  /** Filter by department(s) */
  departmentFilter?: string[];
  /** Filter by user type */
  userTypeFilter?: ('Member' | 'Guest')[];
  /** Exclude specific users by UPN */
  excludeUsers?: string[];
  /** Only sync users in specific Entra groups */
  entraGroupFilter?: string[];
  /** Batch size for processing */
  batchSize: number;
  /** Send notification on completion */
  sendNotification: boolean;
  /** Notification recipients */
  notificationRecipients?: string[];
}

/**
 * Default sync configuration
 */
export const DEFAULT_SYNC_CONFIG: ISyncConfig = {
  includeDisabledUsers: true,
  updateExisting: true,
  deactivateMissing: false,
  userTypeFilter: ['Member'],
  batchSize: 50,
  sendNotification: false
};

/**
 * Sync schedule configuration
 */
export interface ISyncSchedule {
  /** Schedule enabled */
  enabled: boolean;
  /** Frequency type */
  frequency: 'Hourly' | 'Daily' | 'Weekly';
  /** Time of day (for Daily/Weekly) - 24hr format "HH:MM" */
  timeOfDay?: string;
  /** Day of week (for Weekly) - 0=Sunday, 1=Monday, etc. */
  dayOfWeek?: number;
  /** Last run timestamp */
  lastRun?: Date;
  /** Next scheduled run */
  nextRun?: Date;
}

/**
 * Field mapping from Entra to JML Employee
 */
export interface IFieldMapping {
  /** Entra field name */
  entraField: keyof IEntraUser;
  /** JML Employee field name */
  jmlField: keyof IJMLEmployee;
  /** Transform function name (optional) */
  transform?: string;
  /** Whether this field should be synced */
  enabled: boolean;
}

/**
 * Default field mappings
 */
export const DEFAULT_FIELD_MAPPINGS: IFieldMapping[] = [
  { entraField: 'displayName', jmlField: 'Title', enabled: true },
  { entraField: 'givenName', jmlField: 'FirstName', enabled: true },
  { entraField: 'surname', jmlField: 'LastName', enabled: true },
  { entraField: 'mail', jmlField: 'Email', enabled: true },
  { entraField: 'jobTitle', jmlField: 'JobTitle', enabled: true },
  { entraField: 'department', jmlField: 'Department', enabled: true },
  { entraField: 'officeLocation', jmlField: 'Location', enabled: true },
  { entraField: 'mobilePhone', jmlField: 'MobilePhone', enabled: true },
  { entraField: 'employeeId', jmlField: 'EmployeeNumber', enabled: true },
  { entraField: 'id', jmlField: 'EntraObjectId', enabled: true }
];
