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

  /**
   * Shared email shell — Forest Teal branded, table-based layout for Outlook compatibility
   */
  private buildEmailShell(opts: {
    headerGradient: string; headerIcon: string; headerTitle: string;
    bodyBg?: string; content: string; footerText: string;
    ctaUrl: string; ctaLabel: string; ctaColor: string;
  }): string {
    const bg = opts.bodyBg || '#ffffff';
    return `<!DOCTYPE html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1.0"><title>${opts.headerTitle}</title></head>
<body style="margin:0;padding:0;background:#f0fdfa;font-family:'Segoe UI',Tahoma,Arial,sans-serif;">
<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#f0fdfa;padding:24px 0;">
<tr><td align="center">
<table role="presentation" width="600" cellpadding="0" cellspacing="0" style="max-width:600px;width:100%;background:${bg};border-radius:12px;box-shadow:0 4px 24px rgba(13,148,136,0.10);overflow:hidden;">
  <tr><td style="background:${opts.headerGradient};padding:28px 32px;text-align:center;">
    <div style="font-size:32px;margin-bottom:8px;">${opts.headerIcon}</div>
    <div style="font-size:22px;font-weight:700;color:#ffffff;letter-spacing:0.3px;">${opts.headerTitle}</div>
    <div style="width:40px;height:2px;background:rgba(255,255,255,0.5);margin:12px auto 4px;border-radius:1px;"></div>
    <div style="font-size:12px;color:rgba(255,255,255,0.85);letter-spacing:0.5px;">DWx Policy Manager</div>
  </td></tr>
  <tr><td style="padding:28px 32px;">
    ${opts.content}
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="margin-top:24px;">
    <tr><td align="center">
      <a href="${opts.ctaUrl}" style="display:inline-block;background:${opts.ctaColor};color:#ffffff;padding:12px 32px;border-radius:8px;text-decoration:none;font-weight:600;font-size:15px;letter-spacing:0.3px;">${opts.ctaLabel}</a>
    </td></tr></table>
  </td></tr>
  <tr><td style="background:#f8fafc;padding:16px 32px;text-align:center;border-top:1px solid #e2e8f0;">
    <div style="font-size:12px;color:#64748b;">${opts.footerText}</div>
    <div style="font-size:11px;color:#94a3b8;margin-top:4px;">First Digital &mdash; DWx Policy Manager</div>
  </td></tr>
</table>
</td></tr></table></body></html>`;
  }

  private buildTaskAssignmentEmail(
    taskTitle: string,
    taskDescription: string,
    dueDate: Date | undefined,
    processId: number,
    siteUrl: string
  ): string {
    const taskUrl = `${siteUrl}/SitePages/MyTasks.aspx`;
    const content = `
    <p style="font-size:15px;color:#334155;margin:0 0 16px;">You have been assigned a new task that requires your attention.</p>
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#f0fdfa;border:1px solid #99f6e4;border-radius:8px;border-left:4px solid #0d9488;">
    <tr><td style="padding:16px 20px;">
      <div style="font-size:17px;font-weight:600;color:#0f172a;margin-bottom:6px;">${taskTitle}</div>
      ${taskDescription ? `<div style="font-size:14px;color:#475569;margin-bottom:8px;">${taskDescription}</div>` : ''}
      <div style="font-size:13px;color:#64748b;">
        Process ID: <strong>#${processId}</strong>
        ${dueDate ? `<br>Due: <strong>${dueDate.toLocaleDateString()}</strong>` : ''}
      </div>
    </td></tr></table>`;

    return this.buildEmailShell({
      headerGradient: 'linear-gradient(135deg, #0d9488 0%, #0f766e 100%)',
      headerIcon: '\u{1F4CB}',
      headerTitle: 'New Task Assigned',
      content,
      footerText: 'This is an automated notification from the DWx Workflow System.',
      ctaUrl: taskUrl,
      ctaLabel: 'View My Tasks',
      ctaColor: '#0d9488'
    });
  }

  private buildApprovalRequestEmail(
    approvalTitle: string,
    processId: number,
    dueDate: Date | undefined,
    siteUrl: string
  ): string {
    const approvalUrl = `${siteUrl}/SitePages/ApprovalCenter.aspx`;
    const content = `
    <p style="font-size:15px;color:#334155;margin:0 0 16px;">Your approval is required for the following request.</p>
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#fffbeb;border:1px solid #fde68a;border-radius:8px;border-left:4px solid #d97706;">
    <tr><td style="padding:16px 20px;">
      <div style="font-size:17px;font-weight:600;color:#0f172a;margin-bottom:6px;">${approvalTitle}</div>
      <div style="font-size:13px;color:#64748b;">
        Process ID: <strong>#${processId}</strong>
        ${dueDate ? `<br>Please respond by: <strong>${dueDate.toLocaleDateString()}</strong>` : ''}
      </div>
    </td></tr></table>`;

    return this.buildEmailShell({
      headerGradient: 'linear-gradient(135deg, #d97706 0%, #b45309 100%)',
      headerIcon: '\u{2705}',
      headerTitle: 'Approval Required',
      content,
      footerText: 'This is an automated notification from the DWx Workflow System.',
      ctaUrl: approvalUrl,
      ctaLabel: 'Review & Respond',
      ctaColor: '#d97706'
    });
  }

  private buildSLABreachEmail(
    stepName: string,
    hoursOverdue: number,
    processId: number,
    siteUrl: string
  ): string {
    const processUrl = `${siteUrl}/SitePages/ProcessDetails.aspx?processId=${processId}`;
    const content = `
    <p style="font-size:15px;color:#334155;margin:0 0 16px;">A workflow step has breached its SLA and requires <strong>immediate attention</strong>.</p>
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#fef2f2;border:1px solid #fecaca;border-radius:8px;border-left:4px solid #dc2626;">
    <tr><td style="padding:16px 20px;">
      <div style="font-size:17px;font-weight:600;color:#991b1b;margin-bottom:6px;">${stepName}</div>
      <div style="font-size:14px;color:#7f1d1d;">
        <strong>${hoursOverdue} hours overdue</strong><br>
        <span style="color:#64748b;">Process ID: #${processId}</span>
      </div>
    </td></tr></table>
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#fef2f2;border:1px solid #fecaca;border-radius:8px;margin-top:12px;">
    <tr><td style="padding:12px 16px;font-size:13px;color:#991b1b;">
      <strong>Action Required:</strong> Please investigate and resolve this SLA breach immediately to avoid compliance escalation.
    </td></tr></table>`;

    return this.buildEmailShell({
      headerGradient: 'linear-gradient(135deg, #991b1b 0%, #7f1d1d 100%)',
      headerIcon: '\u{1F6A8}',
      headerTitle: 'SLA BREACH ALERT',
      bodyBg: '#ffffff',
      content,
      footerText: 'This is an urgent automated notification from the DWx Workflow System.',
      ctaUrl: processUrl,
      ctaLabel: 'View Process Details',
      ctaColor: '#dc2626'
    });
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
