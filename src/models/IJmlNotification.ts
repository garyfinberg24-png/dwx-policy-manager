// PM_Notifications List Model

import { IBaseListItem, NotificationType, NotificationStatus, Priority, IUser } from './ICommon';

/**
 * Confirmation status for notifications that require acknowledgment
 * INTEGRATION FIX P5: IT Admin notification confirmation tracking
 */
export enum NotificationConfirmationStatus {
  NotRequired = 'NotRequired',
  Pending = 'Pending',
  Confirmed = 'Confirmed',
  Expired = 'Expired'
}

export interface IJmlNotification extends IBaseListItem {
  // Notification Details (Title is the Subject)
  NotificationType: NotificationType;
  MessageBody: string;
  Priority: Priority;

  // Recipient
  RecipientId: number;
  Recipient?: IUser;
  RecipientEmail?: string;

  // Context
  ProcessId?: string;
  TaskId?: string;

  // Status
  Status: NotificationStatus;
  SentDate?: Date;
  ErrorMessage?: string;
  RetryCount?: number;

  // Template
  TemplateUsed?: string;

  // Delivery Details
  DeliveryMethod?: string;
  TeamsMessageId?: string;
  EmailMessageId?: string;

  // Scheduling
  ScheduledDate?: Date;
  ExpiryDate?: Date;

  // INTEGRATION FIX P5: IT Admin Confirmation Tracking
  RequiresConfirmation?: boolean;
  ConfirmationStatus?: NotificationConfirmationStatus;
  ConfirmedAt?: Date;
  ConfirmedById?: number;
  ConfirmedBy?: IUser;
  ConfirmationToken?: string;  // Unique token for confirmation links
  ConfirmationExpiresAt?: Date;
  ConfirmationNotes?: string;

  // Additional
  CustomData?: string; // JSON for additional fields
}

// For notification queue management
export interface IJmlNotificationQueue {
  Id: number;
  Title: string;
  NotificationType: NotificationType;
  Recipient: string;
  Status: NotificationStatus;
  Priority: Priority;
  ScheduledDate?: Date;
  RetryCount?: number;
}

// Notification templates
export interface INotificationTemplate {
  name: string;
  subject: string;
  body: string;
  type: NotificationType;
  variables: string[]; // e.g., ["EmployeeName", "ProcessType", "DueDate"]
}
