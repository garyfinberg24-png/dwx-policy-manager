// JML_TaskTimeEntries List Model
// Tracks time spent on task assignments for analytics and SLA monitoring

import { IBaseListItem, IUser } from './ICommon';

export interface IJmlTaskTimeEntry extends IBaseListItem {
  // Parent Task Reference
  TaskAssignmentId: number;
  TaskAssignment?: {
    Id: number;
    Title: string;
  };

  // User Information
  UserId: number;
  User?: IUser;

  // Time Tracking
  StartTime: Date;
  EndTime?: Date;
  HoursLogged: number;
  IsActive?: boolean; // Currently running timer

  // Work Classification
  WorkType: WorkType;
  ActivityDescription?: string;

  // Metadata
  IsBillable?: boolean;
  Notes?: string;
}

export enum WorkType {
  Active = 'Active',           // Actively working on task
  Research = 'Research',       // Researching/learning
  Waiting = 'Waiting',         // Waiting on dependencies/approvals
  Review = 'Review',           // Reviewing documents/code
  Meeting = 'Meeting',         // Task-related meetings
  Rework = 'Rework'           // Fixing issues/revisions
}

// View model for time tracking display
export interface ITaskTimeEntryView extends IJmlTaskTimeEntry {
  Duration?: string; // Formatted duration (e.g., "2h 30m")
  CanEdit: boolean;
  CanDelete: boolean;
}

// Form model for creating/editing time entries
export interface ITaskTimeEntryForm {
  TaskAssignmentId: number;
  StartTime: Date;
  EndTime?: Date;
  HoursLogged?: number;
  WorkType: WorkType;
  ActivityDescription?: string;
  IsBillable?: boolean;
  Notes?: string;
}

// Summary model for time analytics
export interface ITaskTimeSummary {
  TaskAssignmentId: number;
  TotalHours: number;
  EstimatedHours?: number;
  VarianceHours: number;
  PercentComplete: number;
  ByWorkType: {
    [key in WorkType]?: number;
  };
  LastEntry?: Date;
  ActiveEntry?: IJmlTaskTimeEntry;
}
