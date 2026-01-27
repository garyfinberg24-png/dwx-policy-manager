// JML_TaskEscalations List Model
// Tracks task escalations and notification rules

import { IBaseListItem, IUser } from './ICommon';

export interface IJmlTaskEscalationRule extends IBaseListItem {
  // Task Template Reference
  TaskId?: number;
  Task?: {
    Id: number;
    Title: string;
    TaskCode: string;
  };

  // Global rule (applies to all tasks)
  IsGlobalRule?: boolean;
  AppliesToCategoryFilter?: string; // JSON array of categories
  ApplicesToDepartment?: string;

  // Escalation Trigger
  EscalationTrigger: EscalationTrigger;
  TriggerValue: number; // Hours or days depending on trigger

  // Notification Settings
  NotifyRoles?: string; // JSON array: ['Manager', 'ProcessOwner', 'HR']
  NotifySpecificUsers?: string; // JSON array of user IDs
  EscalationLevel: number; // 1, 2, 3 for progressive escalation

  // Actions
  AutoReassign?: boolean;
  ReassignToRole?: string;
  AutoChangeStatus?: boolean;
  NewStatus?: string;

  // Enabled/Disabled
  IsActive?: boolean;
}

export enum EscalationTrigger {
  OverdueBy = 'OverdueBy',                    // X hours/days past due date
  NotStartedAfter = 'NotStartedAfter',        // X hours/days after assignment
  StuckInStatus = 'StuckInStatus',            // No update in X hours/days
  ApproachingDue = 'ApproachingDue',          // X hours/days before due date
  HighPriorityOverdue = 'HighPriorityOverdue' // High/Critical priority overdue
}

// Escalation Log
export interface IJmlTaskEscalationLog extends IBaseListItem {
  // Task Reference
  TaskAssignmentId: number;
  TaskAssignment?: {
    Id: number;
    Title: string;
  };

  // Escalation Details
  EscalationRuleId: number;
  EscalationRule?: IJmlTaskEscalationRule;
  EscalationTrigger: EscalationTrigger;
  EscalationLevel: number;

  // Notification Details
  NotifiedUsers?: string; // JSON array of user IDs
  NotificationSentDate: Date;
  NotificationMethod: 'Email' | 'Teams' | 'Both';

  // Actions Taken
  WasReassigned?: boolean;
  ReassignedToId?: number;
  ReassignedTo?: IUser;
  StatusChanged?: boolean;
  PreviousStatus?: string;
  NewStatus?: string;

  // Resolution
  IsResolved?: boolean;
  ResolvedDate?: Date;
  ResolutionNotes?: string;
}

// Notification Queue Item
export interface ITaskNotificationQueueItem {
  TaskAssignmentId: number;
  NotificationType: TaskNotificationType;
  ScheduledFor: Date;
  Priority: 'Low' | 'Normal' | 'High' | 'Urgent';
  Recipients: number[]; // User IDs
  Message: string;
  IsProcessed?: boolean;
}

export enum TaskNotificationType {
  Assigned = 'Assigned',           // New task assigned to you
  Reminder = 'Reminder',           // Upcoming due date
  Overdue = 'Overdue',            // Task is overdue
  Escalation = 'Escalation',      // Escalated to manager
  Reassigned = 'Reassigned',      // Task reassigned to you
  Mentioned = 'Mentioned',         // Mentioned in comment
  StatusChange = 'StatusChange',   // Task status changed
  Completed = 'Completed',         // Task completed
  SLAWarning = 'SLAWarning',      // SLA approaching breach threshold
  SLABreach = 'SLABreach'         // SLA has been breached
}
