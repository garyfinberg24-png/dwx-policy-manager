// @ts-nocheck
/**
 * EmailQueueService
 * Queue-based email service for background operations
 *
 * This service queues emails to a SharePoint list (PM_EmailQueue) which can be
 * processed by:
 * 1. Power Automate flow (recommended for production)
 * 2. A scheduled Azure Function
 * 3. Manual processing by an admin
 *
 * This approach works without user context, making it suitable for:
 * - Workflow background tasks
 * - Scheduled reminders
 * - System notifications
 * - SLA breach alerts
 */

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';

import { logger } from './LoggingService';

// ============================================================================
// INTERFACES
// ============================================================================

export interface IEmailQueueItem {
  Id?: number;
  Title: string;
  To: string;           // Comma-separated email addresses
  CC?: string;          // Comma-separated CC addresses
  Subject: string;
  Body: string;         // HTML body
  Priority: EmailPriority;
  Status: EmailQueueStatus;
  // Tracking
  QueuedAt: string;
  SentAt?: string;
  AttemptCount: number;
  LastAttemptAt?: string;
  ErrorMessage?: string;
  // Context
  ProcessId?: number;
  WorkflowInstanceId?: number;
  NotificationType?: string;
  SourceSystem: string;
}

export enum EmailPriority {
  Low = 'Low',
  Normal = 'Normal',
  High = 'High',
  Urgent = 'Urgent'
}

export enum EmailQueueStatus {
  Queued = 'Queued',
  Processing = 'Processing',
  Sent = 'Sent',
  Failed = 'Failed',
  Cancelled = 'Cancelled'
}

export interface IQueueEmailOptions {
  to: string[];
  cc?: string[];
  subject: string;
  htmlBody: string;
  priority?: EmailPriority;
  processId?: number;
  workflowInstanceId?: number;
  notificationType?: string;
}

export interface IQueueResult {
  success: boolean;
  queueItemId?: number;
  error?: string;
}

export interface IBatchQueueResult {
  success: boolean;
  queued: number;
  failed: number;
  queueItemIds: number[];
  errors: string[];
}

// ============================================================================
// EMAIL QUEUE SERVICE
// ============================================================================

export class EmailQueueService {
  private sp: SPFI;
  private readonly listName = 'PM_EmailQueue';
  private readonly sourceSystem = 'JML-Workflow';

  constructor(sp: SPFI) {
    this.sp = sp;
  }

  // ============================================================================
  // PUBLIC METHODS - QUEUEING
  // ============================================================================

  /**
   * Queue a single email for sending
   */
  public async queueEmail(options: IQueueEmailOptions): Promise<IQueueResult> {
    try {
      if (!options.to || options.to.length === 0) {
        return { success: false, error: 'No recipients specified' };
      }

      const queueItem: Partial<IEmailQueueItem> = {
        Title: options.subject.substring(0, 255), // SharePoint Title limit
        To: options.to.join(';'),
        CC: options.cc?.join(';'),
        Subject: options.subject,
        Body: options.htmlBody,
        Priority: options.priority || EmailPriority.Normal,
        Status: EmailQueueStatus.Queued,
        QueuedAt: new Date().toISOString(),
        AttemptCount: 0,
        ProcessId: options.processId,
        WorkflowInstanceId: options.workflowInstanceId,
        NotificationType: options.notificationType,
        SourceSystem: this.sourceSystem
      };

      const result = await this.sp.web.lists
        .getByTitle(this.listName)
        .items.add(queueItem);

      logger.info('EmailQueueService', `Email queued successfully: ${result.data.Id}`, {
        to: options.to,
        subject: options.subject
      });

      return {
        success: true,
        queueItemId: result.data.Id
      };

    } catch (error) {
      const errorMsg = error instanceof Error ? error.message : 'Unknown error';
      logger.error('EmailQueueService', 'Failed to queue email', error);
      return {
        success: false,
        error: errorMsg
      };
    }
  }

  /**
   * Queue multiple emails for batch sending
   */
  public async queueBatchEmails(emails: IQueueEmailOptions[]): Promise<IBatchQueueResult> {
    const result: IBatchQueueResult = {
      success: true,
      queued: 0,
      failed: 0,
      queueItemIds: [],
      errors: []
    };

    for (const email of emails) {
      const queueResult = await this.queueEmail(email);
      if (queueResult.success && queueResult.queueItemId) {
        result.queued++;
        result.queueItemIds.push(queueResult.queueItemId);
      } else {
        result.failed++;
        result.errors.push(queueResult.error || 'Unknown error');
      }
    }

    result.success = result.failed === 0;
    return result;
  }

  /**
   * Queue email for each recipient individually (for personalization)
   */
  public async queueIndividualEmails(
    recipients: Array<{ email: string; name?: string }>,
    subjectTemplate: string,
    bodyTemplate: string,
    options?: Partial<IQueueEmailOptions>
  ): Promise<IBatchQueueResult> {
    const emails: IQueueEmailOptions[] = recipients.map(recipient => ({
      to: [recipient.email],
      subject: this.replaceTokens(subjectTemplate, { recipientName: recipient.name || recipient.email }),
      htmlBody: this.replaceTokens(bodyTemplate, { recipientName: recipient.name || recipient.email }),
      priority: options?.priority,
      processId: options?.processId,
      workflowInstanceId: options?.workflowInstanceId,
      notificationType: options?.notificationType
    }));

    return this.queueBatchEmails(emails);
  }

  // ============================================================================
  // PUBLIC METHODS - QUEUE MANAGEMENT
  // ============================================================================

  /**
   * Get pending emails from queue
   */
  public async getPendingEmails(limit: number = 50): Promise<IEmailQueueItem[]> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(this.listName)
        .items
        .filter(`Status eq 'Queued'`)
        .orderBy('Priority', false) // Urgent first
        .orderBy('QueuedAt', true)  // Oldest first
        .top(limit)();

      return items as IEmailQueueItem[];
    } catch (error) {
      logger.error('EmailQueueService', 'Failed to get pending emails', error);
      return [];
    }
  }

  /**
   * Mark email as sent
   */
  public async markAsSent(queueItemId: number): Promise<boolean> {
    try {
      await this.sp.web.lists
        .getByTitle(this.listName)
        .items.getById(queueItemId)
        .update({
          Status: EmailQueueStatus.Sent,
          SentAt: new Date().toISOString()
        });
      return true;
    } catch (error) {
      logger.error('EmailQueueService', `Failed to mark email ${queueItemId} as sent`, error);
      return false;
    }
  }

  /**
   * Mark email as failed
   */
  public async markAsFailed(queueItemId: number, errorMessage: string): Promise<boolean> {
    try {
      const item = await this.sp.web.lists
        .getByTitle(this.listName)
        .items.getById(queueItemId)
        .select('AttemptCount')();

      const attemptCount = (item.AttemptCount || 0) + 1;
      const maxAttempts = 3;

      await this.sp.web.lists
        .getByTitle(this.listName)
        .items.getById(queueItemId)
        .update({
          Status: attemptCount >= maxAttempts ? EmailQueueStatus.Failed : EmailQueueStatus.Queued,
          AttemptCount: attemptCount,
          LastAttemptAt: new Date().toISOString(),
          ErrorMessage: errorMessage
        });

      return true;
    } catch (error) {
      logger.error('EmailQueueService', `Failed to mark email ${queueItemId} as failed`, error);
      return false;
    }
  }

  /**
   * Cancel a queued email
   */
  public async cancelEmail(queueItemId: number): Promise<boolean> {
    try {
      await this.sp.web.lists
        .getByTitle(this.listName)
        .items.getById(queueItemId)
        .update({
          Status: EmailQueueStatus.Cancelled
        });
      return true;
    } catch (error) {
      logger.error('EmailQueueService', `Failed to cancel email ${queueItemId}`, error);
      return false;
    }
  }

  /**
   * Retry a failed email
   */
  public async retryEmail(queueItemId: number): Promise<boolean> {
    try {
      await this.sp.web.lists
        .getByTitle(this.listName)
        .items.getById(queueItemId)
        .update({
          Status: EmailQueueStatus.Queued,
          ErrorMessage: null
        });
      return true;
    } catch (error) {
      logger.error('EmailQueueService', `Failed to retry email ${queueItemId}`, error);
      return false;
    }
  }

  /**
   * Get queue statistics
   */
  public async getQueueStats(): Promise<{
    queued: number;
    processing: number;
    sent: number;
    failed: number;
    total: number;
  }> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(this.listName)
        .items
        .select('Status')
        .top(5000)();

      const stats = {
        queued: 0,
        processing: 0,
        sent: 0,
        failed: 0,
        total: items.length
      };

      items.forEach((item: { Status: string }) => {
        switch (item.Status) {
          case EmailQueueStatus.Queued:
            stats.queued++;
            break;
          case EmailQueueStatus.Processing:
            stats.processing++;
            break;
          case EmailQueueStatus.Sent:
            stats.sent++;
            break;
          case EmailQueueStatus.Failed:
            stats.failed++;
            break;
        }
      });

      return stats;
    } catch (error) {
      logger.error('EmailQueueService', 'Failed to get queue stats', error);
      return { queued: 0, processing: 0, sent: 0, failed: 0, total: 0 };
    }
  }

  // ============================================================================
  // TEMPLATE METHODS
  // ============================================================================

  /**
   * Queue a workflow notification email
   */
  public async queueWorkflowNotification(
    recipients: string[],
    notificationType: string,
    subject: string,
    htmlBody: string,
    workflowInstanceId: number,
    processId: number,
    priority: EmailPriority = EmailPriority.Normal
  ): Promise<IQueueResult> {
    return this.queueEmail({
      to: recipients,
      subject,
      htmlBody,
      priority,
      workflowInstanceId,
      processId,
      notificationType
    });
  }

  /**
   * Queue a task assignment notification
   */
  public async queueTaskAssignmentNotification(
    assigneeEmail: string,
    taskTitle: string,
    taskDescription: string,
    dueDate: Date | undefined,
    processId: number,
    siteUrl: string
  ): Promise<IQueueResult> {
    const subject = `Task Assigned: ${taskTitle}`;
    const htmlBody = this.buildTaskAssignmentEmail(taskTitle, taskDescription, dueDate, processId, siteUrl);

    return this.queueEmail({
      to: [assigneeEmail],
      subject,
      htmlBody,
      priority: EmailPriority.Normal,
      processId,
      notificationType: 'TaskAssigned'
    });
  }

  /**
   * Queue an approval request notification
   */
  public async queueApprovalRequestNotification(
    approverEmail: string,
    approvalTitle: string,
    processId: number,
    dueDate: Date | undefined,
    siteUrl: string
  ): Promise<IQueueResult> {
    const subject = `Approval Required: ${approvalTitle}`;
    const htmlBody = this.buildApprovalRequestEmail(approvalTitle, processId, dueDate, siteUrl);

    return this.queueEmail({
      to: [approverEmail],
      subject,
      htmlBody,
      priority: EmailPriority.High,
      processId,
      notificationType: 'ApprovalRequested'
    });
  }

  /**
   * Queue an SLA breach notification
   */
  public async queueSLABreachNotification(
    recipients: string[],
    stepName: string,
    hoursOverdue: number,
    processId: number,
    workflowInstanceId: number,
    siteUrl: string
  ): Promise<IQueueResult> {
    const subject = `SLA BREACH: ${stepName} is ${hoursOverdue} hours overdue`;
    const htmlBody = this.buildSLABreachEmail(stepName, hoursOverdue, processId, siteUrl);

    return this.queueEmail({
      to: recipients,
      subject,
      htmlBody,
      priority: EmailPriority.Urgent,
      processId,
      workflowInstanceId,
      notificationType: 'SLABreach'
    });
  }

  // ============================================================================
  // PRIVATE METHODS - EMAIL TEMPLATES
  // ============================================================================

  private buildTaskAssignmentEmail(
    taskTitle: string,
    taskDescription: string,
    dueDate: Date | undefined,
    processId: number,
    siteUrl: string
  ): string {
    const taskUrl = `${siteUrl}/SitePages/MyTasks.aspx`;

    return `
<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <style>
    body { font-family: 'Segoe UI', sans-serif; margin: 0; padding: 0; background: #f3f2f1; }
    .container { max-width: 600px; margin: 0 auto; padding: 20px; }
    .card { background: #fff; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); overflow: hidden; }
    .header { background: #e7f3ff; border-left: 4px solid #0078d4; padding: 16px 20px; }
    .header h1 { margin: 0; font-size: 18px; color: #004578; }
    .content { padding: 20px; }
    .task-card { background: #faf9f8; border: 1px solid #edebe9; border-radius: 4px; padding: 16px; margin: 16px 0; }
    .task-title { font-size: 16px; font-weight: 600; color: #323130; margin-bottom: 8px; }
    .task-desc { font-size: 14px; color: #605e5c; margin-bottom: 8px; }
    .task-meta { font-size: 13px; color: #605e5c; }
    .button { display: inline-block; background: #0078d4; color: #fff; padding: 10px 24px; border-radius: 4px; text-decoration: none; font-weight: 600; margin-top: 16px; }
    .footer { padding: 16px 20px; background: #faf9f8; text-align: center; font-size: 12px; color: #605e5c; }
  </style>
</head>
<body>
  <div class="container">
    <div class="card">
      <div class="header">
        <h1>New Task Assigned</h1>
      </div>
      <div class="content">
        <p>You have been assigned a new task that requires your attention.</p>
        <div class="task-card">
          <div class="task-title">${taskTitle}</div>
          ${taskDescription ? `<div class="task-desc">${taskDescription}</div>` : ''}
          <div class="task-meta">
            Process ID: #${processId}
            ${dueDate ? `<br>Due: ${dueDate.toLocaleDateString()}` : ''}
          </div>
        </div>
        <a href="${taskUrl}" class="button">View My Tasks</a>
      </div>
      <div class="footer">
        This is an automated notification from the JML Workflow System.
      </div>
    </div>
  </div>
</body>
</html>`;
  }

  private buildApprovalRequestEmail(
    approvalTitle: string,
    processId: number,
    dueDate: Date | undefined,
    siteUrl: string
  ): string {
    const approvalUrl = `${siteUrl}/SitePages/ApprovalCenter.aspx`;

    return `
<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <style>
    body { font-family: 'Segoe UI', sans-serif; margin: 0; padding: 0; background: #f3f2f1; }
    .container { max-width: 600px; margin: 0 auto; padding: 20px; }
    .card { background: #fff; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); overflow: hidden; }
    .header { background: #fff4ce; border-left: 4px solid #ff8c00; padding: 16px 20px; }
    .header h1 { margin: 0; font-size: 18px; color: #8a6914; }
    .content { padding: 20px; }
    .approval-card { background: #faf9f8; border: 1px solid #edebe9; border-radius: 4px; padding: 16px; margin: 16px 0; }
    .approval-title { font-size: 16px; font-weight: 600; color: #323130; margin-bottom: 8px; }
    .approval-meta { font-size: 13px; color: #605e5c; }
    .button { display: inline-block; background: #0078d4; color: #fff; padding: 10px 24px; border-radius: 4px; text-decoration: none; font-weight: 600; margin-top: 16px; }
    .footer { padding: 16px 20px; background: #faf9f8; text-align: center; font-size: 12px; color: #605e5c; }
  </style>
</head>
<body>
  <div class="container">
    <div class="card">
      <div class="header">
        <h1>Approval Required</h1>
      </div>
      <div class="content">
        <p>Your approval is required for the following request.</p>
        <div class="approval-card">
          <div class="approval-title">${approvalTitle}</div>
          <div class="approval-meta">
            Process ID: #${processId}
            ${dueDate ? `<br>Please respond by: ${dueDate.toLocaleDateString()}` : ''}
          </div>
        </div>
        <a href="${approvalUrl}" class="button">Review & Respond</a>
      </div>
      <div class="footer">
        This is an automated notification from the JML Workflow System.
      </div>
    </div>
  </div>
</body>
</html>`;
  }

  private buildSLABreachEmail(
    stepName: string,
    hoursOverdue: number,
    processId: number,
    siteUrl: string
  ): string {
    const processUrl = `${siteUrl}/SitePages/ProcessDetails.aspx?processId=${processId}`;

    return `
<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <style>
    body { font-family: 'Segoe UI', sans-serif; margin: 0; padding: 0; background: #f3f2f1; }
    .container { max-width: 600px; margin: 0 auto; padding: 20px; }
    .card { background: #fff; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); overflow: hidden; }
    .header { background: #fde7e9; border-left: 4px solid #d13438; padding: 16px 20px; }
    .header h1 { margin: 0; font-size: 18px; color: #a80000; }
    .content { padding: 20px; }
    .alert-card { background: #fde7e9; border: 1px solid #d13438; border-radius: 4px; padding: 16px; margin: 16px 0; }
    .alert-title { font-size: 16px; font-weight: 600; color: #a80000; margin-bottom: 8px; }
    .alert-meta { font-size: 14px; color: #605e5c; }
    .button { display: inline-block; background: #d13438; color: #fff; padding: 10px 24px; border-radius: 4px; text-decoration: none; font-weight: 600; margin-top: 16px; }
    .footer { padding: 16px 20px; background: #faf9f8; text-align: center; font-size: 12px; color: #605e5c; }
  </style>
</head>
<body>
  <div class="container">
    <div class="card">
      <div class="header">
        <h1>SLA BREACH ALERT</h1>
      </div>
      <div class="content">
        <p>A workflow step has breached its SLA and requires immediate attention.</p>
        <div class="alert-card">
          <div class="alert-title">${stepName}</div>
          <div class="alert-meta">
            <strong>${hoursOverdue} hours overdue</strong><br>
            Process ID: #${processId}
          </div>
        </div>
        <a href="${processUrl}" class="button">View Process Details</a>
      </div>
      <div class="footer">
        This is an urgent automated notification from the JML Workflow System.
      </div>
    </div>
  </div>
</body>
</html>`;
  }

  // ============================================================================
  // PRIVATE METHODS - UTILITIES
  // ============================================================================

  private replaceTokens(template: string, tokens: Record<string, string>): string {
    let result = template;
    for (const [key, value] of Object.entries(tokens)) {
      result = result.replace(new RegExp(`{{${key}}}`, 'g'), value);
    }
    return result;
  }
}

export default EmailQueueService;
