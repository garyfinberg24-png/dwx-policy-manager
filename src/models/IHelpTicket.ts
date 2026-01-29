// PM_HelpTickets List Model
// Support ticket tracking and management

import { IBaseListItem, IUser, Priority } from './ICommon';

export type TicketStatus =
  | 'New'
  | 'Open'
  | 'In Progress'
  | 'Waiting for User'
  | 'Waiting for Developer'
  | 'Resolved'
  | 'Closed'
  | 'Cancelled';

export type TicketCategory =
  | 'Bug Report'
  | 'Feature Request'
  | 'How-To Question'
  | 'Performance Issue'
  | 'Access Issue'
  | 'Data Issue'
  | 'Integration Issue'
  | 'General Inquiry'
  | 'Other';

export type SeverityLevel =
  | 'Critical' // System down
  | 'High' // Major functionality broken
  | 'Medium' // Feature not working as expected
  | 'Low'; // Minor issue or enhancement

export interface IHelpTicket extends IBaseListItem {
  // Ticket Info
  Title: string; // Brief description
  TicketNumber?: string; // Auto-generated (e.g., "HELP-2025-001")
  Category: TicketCategory;
  Status: TicketStatus;
  Priority: Priority;
  Severity: SeverityLevel;

  // Details
  Description: string; // Detailed description
  StepsToReproduce?: string;
  ExpectedBehavior?: string;
  ActualBehavior?: string;

  // Assignment
  SubmittedById?: number;
  SubmittedBy?: IUser;
  AssignedToId?: number;
  AssignedTo?: IUser;

  // Environment Info
  WebPartName?: string;
  PageUrl?: string;
  BrowserInfo?: string;
  ErrorMessage?: string;
  ScreenshotUrl?: string;

  // Tracking
  SubmittedDate?: Date;
  FirstResponseDate?: Date;
  ResolvedDate?: Date;
  ClosedDate?: Date;

  // Resolution
  Resolution?: string;
  ResolutionNotes?: string;
  RelatedArticleId?: number; // Link to help article that resolved issue

  // Metrics
  ResponseTime?: number; // Hours until first response
  ResolutionTime?: number; // Hours until resolved
  SatisfactionRating?: number; // 1-5 stars
  FeedbackComments?: string;

  // Related Items
  RelatedTickets?: string; // Comma-separated ticket IDs
  RelatedProcessId?: number;

  // Follow-up
  RequiresFollowUp: boolean;
  FollowUpDate?: Date;
  FollowUpNotes?: string;
}

// For ticket summary/dashboard
export interface IHelpTicketSummary {
  TotalTickets: number;
  OpenTickets: number;
  InProgressTickets: number;
  ResolvedToday: number;
  AvgResponseTime: number;
  AvgResolutionTime: number;
  SatisfactionScore: number;
  ByCategory: { [key: string]: number };
  BySeverity: { [key: string]: number };
  ByStatus: { [key: string]: number };
}
