// Approval Workflow Models
// Multi-level approval processes for JML workflows

import { IUser } from './ICommon';

export enum ApprovalStatus {
  Pending = 'Pending',
  Approved = 'Approved',
  Rejected = 'Rejected',
  Delegated = 'Delegated',
  Escalated = 'Escalated',
  Cancelled = 'Cancelled',
  Skipped = 'Skipped',
  /** GAP FIX: Status for sequential approvers waiting for their turn */
  Queued = 'Queued',
  /** GAP FIX: Status for approvals that have exceeded maximum allowed time */
  Expired = 'Expired'
}

export enum ApprovalType {
  Sequential = 'Sequential',
  Parallel = 'Parallel',
  FirstApprover = 'FirstApprover'
}

export enum EscalationAction {
  Notify = 'Notify',
  AutoApprove = 'AutoApprove',
  AssignToManager = 'AssignToManager',
  AssignToAlternate = 'AssignToAlternate'
}

export interface IJmlApproval {
  Id: number;
  ProcessID: number;
  ProcessTitle: string;
  ProcessType: string;

  // Approval Details
  ApprovalLevel: number;
  ApprovalSequence: number;
  ApprovalType: ApprovalType;
  Status: ApprovalStatus;

  // Approver Information
  ApproverId: number;
  Approver: IUser;
  OriginalApproverId?: number;
  OriginalApprover?: IUser;
  DelegatedById?: number;
  DelegatedBy?: IUser;

  // Request Details
  RequestedDate: Date;
  DueDate: Date;
  CompletedDate?: Date;
  ResponseTime?: number; // in hours

  // Decision Details
  Decision?: ApprovalStatus;
  Comments?: string;
  Notes?: string;
  ReasonRequired: boolean;
  ActualCompletionDate?: Date;

  // Escalation
  IsOverdue: boolean;
  EscalationLevel: number;
  EscalationDate?: Date;
  EscalationAction?: EscalationAction;

  // Workflow Integration
  WorkflowInstanceId?: number;
  WorkflowStepId?: string;
  ApprovalTemplateId?: number;

  // Metadata
  Created: Date;
  Modified: Date;
  ModifiedBy: IUser;
}

export interface IJmlApprovalChain {
  Id: number;
  ProcessID: number;
  ChainName: string;
  ApprovalType: ApprovalType;
  IsActive: boolean;

  // Chain Configuration
  Levels: IJmlApprovalLevel[];
  RequireComments: boolean;
  AllowDelegation: boolean;
  AutoEscalationDays: number;
  EscalationAction: EscalationAction;

  // Status
  CurrentLevel: number;
  OverallStatus: ApprovalStatus;
  StartDate?: Date;
  CompletedDate?: Date;

  Created: Date;
  CreatedBy: IUser;
  Modified: Date;
  ModifiedBy: IUser;
}

export interface IJmlApprovalLevel {
  Level: number;
  ApproverIds: number[];
  Approvers: IUser[];
  ApprovalType: ApprovalType; // Sequential or Parallel
  DueDays: number;
  ReasonRequired: boolean;
  AllowDelegation: boolean;
  EscalateToManagerOnDelay: boolean;
}

export interface IJmlApprovalHistory {
  Id: number;
  ApprovalId: number;
  ProcessID: number;

  Action: string; // Approved, Rejected, Delegated, Escalated
  PerformedBy: IUser;
  PerformedById: number;
  ActionDate: Date;

  Comments?: string;
  Notes?: string;
  PreviousStatus: ApprovalStatus;
  NewStatus: ApprovalStatus;

  // For Delegation
  DelegatedTo?: IUser;
  DelegatedToId?: number;
  DelegationReason?: string;

  // For Escalation
  EscalatedTo?: IUser;
  EscalatedToId?: number;
  EscalationReason?: string;

  Created: Date;
}

export interface IJmlApprovalDelegation {
  Id: number;
  DelegatedById: number;
  DelegatedBy: IUser;
  DelegatedToId: number;
  DelegatedTo: IUser;

  StartDate: Date;
  EndDate: Date;
  IsActive: boolean;

  Reason?: string;
  ProcessTypes?: string[]; // Empty means all types
  AutoDelegate: boolean;

  Created: Date;
  Modified: Date;
}

export interface IJmlApprovalTemplate {
  Id: number;
  Title: string;
  Description?: string;
  ProcessTypes: string[];

  ApprovalType: ApprovalType;
  Levels: IJmlApprovalLevel[];

  RequireComments: boolean;
  AllowDelegation: boolean;
  AutoEscalationDays: number;
  EscalationAction: EscalationAction;

  IsActive: boolean;
  Created: Date;
  CreatedBy: IUser;
  Modified: Date;
  ModifiedBy: IUser;
}

// Request/Response Models

export interface IApprovalRequest {
  processId: number;
  templateId?: number;
  customChain?: IJmlApprovalLevel[];
  comments?: string;
}

export interface IApprovalDecision {
  approvalId: number;
  decision: ApprovalStatus;
  comments: string;
  notifyNext: boolean;
}

export interface IApprovalDelegationRequest {
  delegateToId: number;
  startDate: Date;
  endDate: Date;
  reason?: string;
  processTypes?: string[];
  autoDelegate: boolean;
}

export interface IApprovalEscalation {
  approvalId: number;
  escalateToId: number;
  reason: string;
  action: EscalationAction;
}

export interface IApprovalSummary {
  totalPending: number;
  totalApproved: number;
  totalRejected: number;
  overdueCount: number;
  avgResponseTime: number;
  myPendingCount: number;
  delegatedCount: number;
}

export interface IApprovalFilters {
  status?: ApprovalStatus[];
  processTypes?: string[];
  approvers?: number[];
  dateFrom?: Date;
  dateTo?: Date;
  isOverdue?: boolean;
  processId?: number;
}

export interface IApprovalNotification {
  approvalId: number;
  recipientId: number;
  notificationType: 'NewApproval' | 'Reminder' | 'Escalation' | 'Delegated' | 'Completed';
  subject: string;
  body: string;
  sendEmail: boolean;
  sendInApp: boolean;
}
