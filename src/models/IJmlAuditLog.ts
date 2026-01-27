// JML_AuditLog List Model

import { IBaseListItem, IUser } from './ICommon';

export interface IJmlAuditLog extends IBaseListItem {
  // Event Information
  EventType: string; // e.g., "Process Created", "Task Completed", "Status Changed"
  EntityType: string; // e.g., "Process", "Task", "Configuration"
  EntityId?: number;

  // User Information
  UserId?: number;
  User?: IUser;
  UserEmail?: string;

  // Change Details
  Action: string; // e.g., "Create", "Update", "Delete", "Approve"
  FieldChanged?: string;
  OldValue?: string;
  NewValue?: string;

  // Context
  ProcessId?: number;
  TaskId?: number;
  Description?: string;

  // Metadata
  Timestamp: Date;
  IPAddress?: string;
  UserAgent?: string;
  SessionId?: string;

  // Additional Data
  AdditionalData?: string; // JSON for complex audit data
  Severity?: string; // "Info", "Warning", "Error"
}

// For audit trail display
export interface IJmlAuditLogSummary {
  Id: number;
  Timestamp: Date;
  User: string;
  EventType: string;
  Description: string;
  EntityType: string;
  EntityId?: number;
}

// Audit event types (constants)
export class AuditEventTypes {
  // Process Events
  static readonly PROCESS_CREATED = 'Process Created';
  static readonly PROCESS_UPDATED = 'Process Updated';
  static readonly PROCESS_COMPLETED = 'Process Completed';
  static readonly PROCESS_CANCELLED = 'Process Cancelled';
  static readonly PROCESS_STATUS_CHANGED = 'Process Status Changed';

  // Task Events
  static readonly TASK_ASSIGNED = 'Task Assigned';
  static readonly TASK_STARTED = 'Task Started';
  static readonly TASK_COMPLETED = 'Task Completed';
  static readonly TASK_APPROVED = 'Task Approved';
  static readonly TASK_REJECTED = 'Task Rejected';

  // Configuration Events
  static readonly CONFIG_UPDATED = 'Configuration Updated';
  static readonly TEMPLATE_CREATED = 'Template Created';
  static readonly TEMPLATE_UPDATED = 'Template Updated';

  // System Events
  static readonly USER_LOGIN = 'User Login';
  static readonly PERMISSION_CHANGED = 'Permission Changed';
  static readonly BULK_OPERATION = 'Bulk Operation';
}
