// PM_Tasks List Model

import { IBaseListItem, TaskCategory, Priority, IUser } from './ICommon';

export interface IJmlTask extends IBaseListItem {
  // Task Identification
  TaskCode: string; // Unique identifier (e.g., "IT-ACC-001")
  Category: TaskCategory;

  // Task Details
  Description?: string;
  Instructions?: string;
  Department: string;

  // Assignment
  DefaultAssigneeId?: number;
  DefaultAssignee?: IUser;
  AssigneeRole?: string; // e.g., "IT Manager", "HR Coordinator"

  // Timing
  SLAHours?: number; // Service Level Agreement in hours
  EstimatedHours?: number;

  // Requirements
  RequiresApproval: boolean;
  ApproverRole?: string;

  // Dependencies
  DependsOn?: string; // TaskCode of prerequisite task
  BlockingTask?: boolean; // If true, must be completed before dependent tasks

  // Metadata
  IsActive: boolean;
  Priority: Priority;
  Tags?: string;

  // Resources
  DocumentationUrl?: string;
  FormUrl?: string;
  SystemUrl?: string;
  RelatedLinks?: string; // JSON string or comma-separated links

  // Automation
  AutomationAvailable?: boolean; // Whether task can be automated

  // Usage Stats
  TimesAssigned?: number;
  AverageCompletionTime?: number; // In hours
}

// For task library/picker
export interface IJmlTaskOption {
  Id: number;
  Title: string;
  TaskCode: string;
  Category: TaskCategory;
  Department: string;
  EstimatedHours?: number;
}
