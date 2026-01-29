// Support Ticket Model
// Interface for support ticket system

import { IBaseListItem, IUser } from './ICommon';

/**
 * Support Ticket
 */
export interface ISupportTicket extends IBaseListItem {
  // Ticket Information
  Description: string; // Detailed issue description
  Category: TicketCategory;
  Priority: TicketPriority;
  Status: TicketStatus;

  // User Information
  SubmittedById: number;
  SubmittedBy?: IUser;
  SubmittedDate: Date;

  // Assignment
  AssignedToId?: number;
  AssignedTo?: IUser;
  AssignedDate?: Date;

  // Resolution
  ResolutionNotes?: string;
  ResolvedDate?: Date;
  ClosedDate?: Date;

  // Related Items
  RelatedProcessId?: number; // Link to PM_Process
  RelatedTaskId?: number; // Link to PM_Task
  Attachments?: string[]; // File URLs

  // Metrics
  ResponseTime?: number; // Hours to first response
  ResolutionTime?: number; // Hours to resolution

  // Feedback
  SatisfactionRating?: number; // 1-5 stars
  SatisfactionFeedback?: string;
}

/**
 * Ticket Categories
 */
export enum TicketCategory {
  Technical = 'Technical Issue',
  ProcessQuestion = 'Process Question',
  AccessRequest = 'Access Request',
  FeatureRequest = 'Feature Request',
  BugReport = 'Bug Report',
  Training = 'Training Request',
  Other = 'Other'
}

/**
 * Ticket Priority
 */
export enum TicketPriority {
  Low = 'Low',
  Medium = 'Medium',
  High = 'High',
  Critical = 'Critical'
}

/**
 * Ticket Status
 */
export enum TicketStatus {
  New = 'New',
  InProgress = 'In Progress',
  WaitingOnUser = 'Waiting on User',
  Resolved = 'Resolved',
  Closed = 'Closed',
  Cancelled = 'Cancelled'
}

/**
 * Support Ticket Form Data
 */
export interface ISupportTicketForm {
  Title: string;
  Description: string;
  Category: TicketCategory;
  Priority: TicketPriority;
  RelatedProcessId?: number;
  RelatedTaskId?: number;
  Attachments?: File[];
}

/**
 * Ticket Summary for List View
 */
export interface ISupportTicketSummary {
  Id: number;
  Title: string;
  Category: TicketCategory;
  Priority: TicketPriority;
  Status: TicketStatus;
  SubmittedDate: Date;
  SubmittedBy: string; // Display name
  AssignedTo?: string; // Display name
  IsOverdue: boolean;
  DaysSinceSubmitted: number;
}

/**
 * Ticket Statistics
 */
export interface ITicketStatistics {
  total: number;
  new: number;
  inProgress: number;
  resolved: number;
  closed: number;
  averageResolutionTime: number; // Hours
  averageResponseTime: number; // Hours
  satisfactionRating: number; // Average 1-5
  byCategory: Record<TicketCategory, number>;
  byPriority: Record<TicketPriority, number>;
}
