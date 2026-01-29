// PM_Processes List Model

import { IBaseListItem, ProcessType, ProcessStatus, Priority, IUser } from './ICommon';

// Re-export ProcessType so other modules can import it from this file
export { ProcessType } from './ICommon';

export interface IJmlProcess extends IBaseListItem {
  // Core Fields
  ProcessType: ProcessType;
  ProcessStatus: ProcessStatus;

  // Employee Information
  EmployeeName: string;
  EmployeeEmail: string;
  EmployeeID?: string; // Changed from EmployeeId to match SharePoint field
  Department: string;
  JobTitle: string;
  Location: string;

  // Manager Information
  ManagerId?: number;  // For Joiner: new manager; For Mover: new manager
  Manager?: IUser;
  CurrentManagerId?: number;  // Mover: manager before transfer
  CurrentManager?: IUser;
  ProcessOwnerId?: number;
  ProcessOwner?: IUser;

  // Mover-Specific Fields
  PreviousDepartment?: string;  // Department before transfer
  PreviousJobTitle?: string;    // Job title before transfer
  PreviousLocation?: string;    // Location before transfer (CurrentLocation)
  IsLocationChange?: boolean;   // Whether this involves a location change
  TransferReason?: string;      // Reason for internal transfer

  // Leaver-Specific Fields
  LastWorkingDay?: Date;        // Final day of employment
  ResignationType?: 'Voluntary' | 'Involuntary' | 'Retirement' | 'Contract End';
  ResignationReason?: string;   // Why employee is leaving
  ExitInterviewCompleted?: boolean;
  ExitInterviewDate?: Date;
  AssetsReturned?: boolean;
  AssetsReturnedDate?: Date;
  ForwardingEmail?: string;     // Personal email for post-employment contact
  RehireEligible?: boolean;     // Whether employee is eligible for rehire

  // Process Details
  StartDate: Date;
  TargetCompletionDate: Date;
  ActualCompletionDate?: Date;
  Priority: Priority;

  // Template Reference
  ChecklistTemplateID?: string; // Text field placeholder (will be converted to Lookup later)

  // Progress Tracking
  TotalTasks?: number;
  CompletedTasks?: number;
  ProgressPercentage?: number; // Changed from PercentComplete to match SharePoint field
  OverdueTasks?: number;

  // Additional Details
  Comments?: string; // Changed from Notes to match SharePoint field
  BusinessUnit?: string;
  CostCenter?: string;
  ContractType?: string;

  // Workflow
  ApprovalRequired?: boolean;
  ApprovalStatus?: string;
  ApprovedBy?: IUser;
  ApprovedById?: number;
  ApprovedDate?: Date;

  // Metadata
  IsDeleted?: boolean;
  Tags?: string;
  CustomFields?: string; // JSON string for flexibility

  // Calculated/Read-only
  DaysRemaining?: number;
  IsOverdue?: boolean;
  StatusSummary?: string;
}

// View model for dashboard display
export interface IJmlProcessSummary {
  Id: number;
  ProcessType: ProcessType;
  ProcessStatus: ProcessStatus;
  EmployeeName: string;
  Department: string;
  StartDate: Date;
  TargetCompletionDate: Date;
  ProgressPercentage: number; // Changed from PercentComplete
  Priority: Priority;
  IsOverdue: boolean;
  Manager?: IUser;
}

// Form model for creating/editing
export interface IJmlProcessForm {
  ProcessType: ProcessType;
  EmployeeName: string;
  EmployeeEmail: string;
  Department: string;
  JobTitle: string;
  Location: string;
  ManagerId?: number;
  ProcessOwnerId?: number;
  StartDate: Date;
  TargetCompletionDate: Date;
  Priority: Priority;
  ChecklistTemplateID?: string;
  Comments?: string; // Changed from Notes
  BusinessUnit?: string;
  CostCenter?: string;
  ContractType?: string;

  // Mover-specific form fields
  CurrentManagerId?: number;
  PreviousDepartment?: string;
  PreviousJobTitle?: string;
  PreviousLocation?: string;
  IsLocationChange?: boolean;
  TransferReason?: string;

  // Leaver-specific form fields
  LastWorkingDay?: Date;
  ResignationType?: 'Voluntary' | 'Involuntary' | 'Retirement' | 'Contract End';
  ResignationReason?: string;
  ForwardingEmail?: string;
  RehireEligible?: boolean;
}
