// Common types and enums used across the JML solution

export enum ProcessType {
  Joiner = 'Joiner',
  Mover = 'Mover',
  Leaver = 'Leaver'
}

export enum ProcessStatus {
  Draft = 'Draft',
  NotStarted = 'Not Started',
  Pending = 'Pending',
  PendingApproval = 'Pending Approval',
  InProgress = 'In Progress',
  OnHold = 'On Hold',
  Completed = 'Completed',
  Cancelled = 'Cancelled',
  Archived = 'Archived'
}

export enum TaskStatus {
  NotStarted = 'Not Started',
  InProgress = 'In Progress',
  Waiting = 'Waiting',
  Blocked = 'Blocked',
  Completed = 'Completed',
  Cancelled = 'Cancelled',
  Skipped = 'Skipped'
}

export enum Priority {
  Low = 'Low',
  Medium = 'Medium',
  High = 'High',
  Critical = 'Critical'
}

export enum TaskCategory {
  ITAccess = 'IT - Access',
  ITEquipment = 'IT - Equipment',
  ITSoftware = 'IT - Software',
  HRDocumentation = 'HR - Documentation',
  HROnboarding = 'HR - Onboarding',
  HROffboarding = 'HR - Offboarding',
  FacilitiesAccess = 'Facilities - Access',
  FacilitiesEquipment = 'Facilities - Equipment',
  FinancePayroll = 'Finance - Payroll',
  TrainingOrientation = 'Training - Orientation',
  SecurityCompliance = 'Security - Compliance',
  Other = 'Other'
}

export enum NotificationType {
  Email = 'Email',
  TeamsMessage = 'Teams Message',
  InApp = 'In-App',
  SMS = 'SMS',
  SystemAlert = 'System Alert'
}

export enum NotificationStatus {
  Pending = 'Pending',
  Sent = 'Sent',
  Failed = 'Failed',
  Cancelled = 'Cancelled'
}

// Base interface for all SharePoint list items
export interface IBaseListItem {
  Id?: number;
  Title: string;
  Created?: Date;
  Modified?: Date;
  AuthorId?: number;
  EditorId?: number;
}

// Enhanced base interface with comprehensive audit fields
export interface IAuditableListItem extends IBaseListItem {
  // Standard SharePoint audit
  Author?: IUser;
  Editor?: IUser;

  // Extended audit tracking
  CreatedByName?: string;
  ModifiedByName?: string;

  // Version tracking
  Version?: string;
  VersionNumber?: number;

  // Soft delete support
  IsDeleted?: boolean;
  DeletedDate?: Date;
  DeletedById?: number;
  DeletedBy?: IUser;
  DeletionReason?: string;

  // Change tracking
  LastChangeType?: 'Create' | 'Update' | 'Delete' | 'Restore';
  LastChangeDate?: Date;
  LastChangedById?: number;
  LastChangedBy?: IUser;
  ChangeHistory?: string; // JSON array of change records

  // Data integrity
  Checksum?: string;
  IsSynced?: boolean;
  SyncStatus?: 'Pending' | 'InProgress' | 'Synced' | 'Failed';
  LastSyncDate?: Date;
  SyncErrorMessage?: string;

  // Archive support
  IsArchived?: boolean;
  ArchivedDate?: Date;
  ArchivedById?: number;
  ArchivedBy?: IUser;
  ArchiveReason?: string;
}

// Audit log entry interface
export interface IAuditLogEntry {
  id: string;
  timestamp: Date;
  entityType: string;
  entityId: number;
  action: 'Create' | 'Read' | 'Update' | 'Delete' | 'Restore' | 'Archive' | 'Sync';
  userId: number;
  userName: string;
  userEmail: string;
  previousValue?: Record<string, unknown>;
  newValue?: Record<string, unknown>;
  changedFields?: string[];
  ipAddress?: string;
  userAgent?: string;
  correlationId?: string;
  workflowInstanceId?: number;
  processId?: number;
  notes?: string;
}

// Change tracking for individual fields
export interface IFieldChange {
  fieldName: string;
  previousValue: unknown;
  newValue: unknown;
  changedAt: Date;
  changedBy: number;
}

// Audit summary for entities
export interface IAuditSummary {
  entityType: string;
  entityId: number;
  totalChanges: number;
  firstCreated: Date;
  lastModified: Date;
  createdBy: IUser;
  lastModifiedBy: IUser;
  topChangers: Array<{ userId: number; userName: string; changeCount: number }>;
  changesByType: Record<string, number>;
}

// User interface for people/group columns
export interface IUser {
  Id?: number;
  Title: string;
  EMail?: string;
}

// Extended user with additional properties
export interface IUserExtended extends IUser {
  LoginName?: string;
  Department?: string;
  JobTitle?: string;
  Manager?: IUser;
  ProfileImageUrl?: string;
  IsActive?: boolean;
}
