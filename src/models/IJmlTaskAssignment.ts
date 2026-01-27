// JML_TaskAssignments List Model

import { IBaseListItem, TaskStatus, Priority, IUser } from './ICommon';
import { ApprovalStatus } from './IJmlApproval';

export interface IJmlTaskAssignment extends IBaseListItem {
  // References
  // Note: ProcessID and TaskID are Text fields until Script 09 (Configure-Lookups.ps1) is executed
  // When Text: ProcessID contains the ID as a string
  // When Lookup (after Script 09): ProcessIDId contains the ID, ProcessID contains the expanded object
  ProcessIDId?: number; // Only present after Script 09 converts to Lookup
  ProcessID?: string | { Id: number; Title: string }; // Text or Lookup object (after Script 09)
  TaskIDId?: number; // Only present after Script 09 converts to Lookup
  TaskID?: string | { Id: number; Title: string; TaskCode?: string; Category?: string }; // Text or Lookup object (after Script 09)

  // Assignment
  AssignedToId: number;
  AssignedTo?: IUser;
  AssignedDate?: Date;

  // Timing
  DueDate: Date;
  StartDate?: Date;
  ActualCompletionDate?: Date;

  // Status
  Status: TaskStatus;
  Priority: Priority;
  PercentComplete?: number;

  // Work Details
  ActualHours?: number;
  Notes?: string;
  CompletionNotes?: string;

  // Approval (if required)
  RequiresApproval?: boolean;
  ApprovalStatus?: ApprovalStatus;
  ApproverId?: number;
  Approver?: IUser;
  ApprovedDate?: Date;
  ApprovalComments?: string;

  // Blocking/Dependencies
  IsDependentTask?: boolean;
  DependsOnTaskId?: number;
  IsBlocked?: boolean;
  BlockedReason?: string;

  // Notifications
  ReminderSent?: boolean;
  EscalationSent?: boolean;
  LastReminderDate?: Date;

  // SLA & Escalation
  SLAHours?: number;           // SLA target hours for completion
  EscalationLevel?: number;    // Current escalation level (0 = none, 1+ = escalated)

  // Workflow Integration
  WorkflowInstanceId?: number; // ID of linked workflow instance
  WorkflowStepId?: string;     // ID of workflow step that created this task

  // Calculated Fields
  IsOverdue?: boolean; // =IF(AND([DueDate]<TODAY(),[Status]<>"Completed"),"Yes","No")
  DaysOverdue?: number;
  DaysRemaining?: number;

  // Metadata
  IsDeleted?: boolean;
  CustomData?: string; // JSON for additional fields
}

// For My Tasks view
export interface IJmlMyTaskSummary {
  Id: number;
  Title: string;
  ProcessId: number;
  ProcessEmployeeName: string;
  ProcessType: string;
  Status: TaskStatus;
  Priority: Priority;
  DueDate: Date;
  IsOverdue: boolean;
  PercentComplete?: number;
}

// For Task Board (Kanban)
export interface IJmlTaskBoardItem {
  Id: number;
  Title: string;
  ProcessId: number;
  EmployeeName: string;
  AssignedTo: IUser;
  Status: TaskStatus;
  Priority: Priority;
  DueDate: Date;
  Category: string;
}
