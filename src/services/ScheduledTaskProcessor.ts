// @ts-nocheck
/**
 * ScheduledTaskProcessor
 * Handles scheduled processing of tasks for escalation, SLA enforcement, and notifications
 * Designed to be called by Power Automate HTTP triggers or Azure Functions
 */

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/items/get-all';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { MSGraphClientV3 } from '@microsoft/sp-http';

import {
  IJmlTaskAssignment,
  ITaskNotificationQueueItem,
  TaskNotificationType,
  EscalationTrigger
} from '../models';
import { TaskStatus, Priority } from '../models/ICommon';
import { TaskNotificationService } from './TaskNotificationService';
import { logger } from './LoggingService';

/**
 * Processing result for scheduled jobs
 */
export interface IScheduledProcessingResult {
  success: boolean;
  processedAt: Date;
  duration: number;
  tasksProcessed: number;
  escalationsGenerated: number;
  notificationsSent: number;
  emailsSent: number;
  teamsMessagesSent: number;
  slaBreaches: number;
  slaWarnings: number;
  errors: string[];
}

/**
 * SLA status for a task
 */
export interface ITaskSLAStatus {
  taskId: number;
  taskTitle: string;
  status: 'Healthy' | 'Warning' | 'Breached';
  slaHours: number;
  elapsedHours: number;
  remainingHours: number;
  assigneeId?: number;
  assigneeEmail?: string;
}

/**
 * Configuration for scheduled processing
 */
export interface IScheduledProcessorConfig {
  enableEscalation: boolean;
  enableSLAEnforcement: boolean;
  enableNotifications: boolean;
  slaWarningThresholdPercent: number; // e.g., 80 = warn at 80% of SLA
  maxTasksPerRun: number;
  dryRun: boolean; // If true, don't actually send notifications
}

const DEFAULT_CONFIG: IScheduledProcessorConfig = {
  enableEscalation: true,
  enableSLAEnforcement: true,
  enableNotifications: true,
  slaWarningThresholdPercent: 80,
  maxTasksPerRun: 500,
  dryRun: false
};

export class ScheduledTaskProcessor {
  private sp: SPFI;
  private context: WebPartContext;
  private taskNotificationService: TaskNotificationService;
  private tasksListTitle = 'JML_TaskAssignments';
  private notificationsListTitle = 'JML_Notifications';
  private slaLogListTitle = 'JML_SLALog';

  constructor(sp: SPFI, context: WebPartContext) {
    this.sp = sp;
    this.context = context;
    this.taskNotificationService = new TaskNotificationService(sp);
  }

  // ============================================================================
  // MAIN PROCESSING ENTRY POINT
  // ============================================================================

  /**
   * Run scheduled processing - main entry point for Power Automate/Azure Functions
   * This method should be called on a schedule (e.g., every 15-30 minutes)
   */
  public async runScheduledProcessing(
    config: Partial<IScheduledProcessorConfig> = {}
  ): Promise<IScheduledProcessingResult> {
    const startTime = Date.now();
    const mergedConfig = { ...DEFAULT_CONFIG, ...config };

    const result: IScheduledProcessingResult = {
      success: true,
      processedAt: new Date(),
      duration: 0,
      tasksProcessed: 0,
      escalationsGenerated: 0,
      notificationsSent: 0,
      emailsSent: 0,
      teamsMessagesSent: 0,
      slaBreaches: 0,
      slaWarnings: 0,
      errors: []
    };

    try {
      logger.info('ScheduledTaskProcessor', 'Starting scheduled processing run');

      // Get all incomplete tasks
      const tasks = await this.getIncompleteTasks(mergedConfig.maxTasksPerRun);
      result.tasksProcessed = tasks.length;

      logger.info('ScheduledTaskProcessor', `Processing ${tasks.length} incomplete tasks`);

      // Process escalations
      if (mergedConfig.enableEscalation) {
        const escalationResult = await this.processEscalations(tasks, mergedConfig);
        result.escalationsGenerated = escalationResult.escalationsGenerated;
        result.errors.push(...escalationResult.errors);
      }

      // Process SLA enforcement
      if (mergedConfig.enableSLAEnforcement) {
        const slaResult = await this.processSLAEnforcement(tasks, mergedConfig);
        result.slaBreaches = slaResult.breaches;
        result.slaWarnings = slaResult.warnings;
        result.errors.push(...slaResult.errors);
      }

      // Send queued notifications
      if (mergedConfig.enableNotifications && !mergedConfig.dryRun) {
        const notificationResult = await this.processNotificationQueue();
        result.notificationsSent = notificationResult.sent;
        result.emailsSent = notificationResult.emailsSent;
        result.teamsMessagesSent = notificationResult.teamsMessagesSent;
        result.errors.push(...notificationResult.errors);
      }

      result.success = result.errors.length === 0;

    } catch (error) {
      result.success = false;
      result.errors.push(error instanceof Error ? error.message : 'Unknown error during scheduled processing');
      logger.error('ScheduledTaskProcessor', 'Scheduled processing failed', error);
    }

    result.duration = Date.now() - startTime;
    logger.info('ScheduledTaskProcessor', `Scheduled processing completed in ${result.duration}ms`, result);

    // Log the run result
    await this.logProcessingRun(result);

    return result;
  }

  // ============================================================================
  // TASK RETRIEVAL
  // ============================================================================

  /**
   * Get all incomplete tasks for processing
   */
  private async getIncompleteTasks(maxTasks: number): Promise<IJmlTaskAssignment[]> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(this.tasksListTitle)
        .items.filter(`Status ne '${TaskStatus.Completed}' and Status ne '${TaskStatus.Cancelled}'`)
        .select(
          'Id', 'Title', 'Status', 'Priority', 'DueDate', 'AssignedToId',
          'AssignedDate', 'ProcessIDId', 'TaskIDId', 'Modified', 'Created',
          'SLAHours', 'IsBlocked', 'BlockedReason', 'EscalationLevel',
          'AssignedTo/Id', 'AssignedTo/Title', 'AssignedTo/EMail',
          'TaskID/Id', 'TaskID/Title', 'TaskID/Category',
          'ProcessID/Id', 'ProcessID/Title', 'ProcessID/EmployeeId'
        )
        .expand('AssignedTo', 'TaskID', 'ProcessID')
        .top(maxTasks)();

      return items.map((item: Record<string, unknown>) => this.mapToTaskAssignment(item));
    } catch (error) {
      logger.error('ScheduledTaskProcessor', 'Error fetching incomplete tasks', error);
      throw error;
    }
  }

  /**
   * Map SharePoint item to IJmlTaskAssignment
   */
  private mapToTaskAssignment(item: Record<string, unknown>): IJmlTaskAssignment {
    const assignedTo = item.AssignedTo as Record<string, unknown> | undefined;
    const taskId = item.TaskID as Record<string, unknown> | undefined;

    return {
      Id: item.Id as number,
      Title: item.Title as string,
      Status: item.Status as TaskStatus,
      Priority: item.Priority as Priority,
      DueDate: item.DueDate ? new Date(item.DueDate as string) : undefined,
      AssignedToId: item.AssignedToId as number,
      AssignedTo: assignedTo ? {
        Id: assignedTo.Id as number,
        Title: assignedTo.Title as string,
        EMail: assignedTo.EMail as string
      } : undefined,
      AssignedDate: item.AssignedDate ? new Date(item.AssignedDate as string) : undefined,
      ProcessIDId: item.ProcessIDId as number,
      TaskIDId: item.TaskIDId as number,
      TaskID: taskId as IJmlTaskAssignment['TaskID'],
      Modified: new Date(item.Modified as string),
      Created: new Date(item.Created as string),
      SLAHours: item.SLAHours as number | undefined,
      IsBlocked: item.IsBlocked as boolean,
      BlockedReason: item.BlockedReason as string,
      EscalationLevel: item.EscalationLevel as number
    };
  }

  // ============================================================================
  // ESCALATION PROCESSING
  // ============================================================================

  /**
   * Process task escalations
   */
  private async processEscalations(
    tasks: IJmlTaskAssignment[],
    config: IScheduledProcessorConfig
  ): Promise<{ escalationsGenerated: number; errors: string[] }> {
    let escalationsGenerated = 0;
    const errors: string[] = [];

    for (const task of tasks) {
      try {
        const notifications = await this.taskNotificationService.evaluateTaskEscalation(task);

        if (notifications.length > 0) {
          escalationsGenerated += notifications.length;

          // Queue notifications for sending
          if (!config.dryRun) {
            await this.queueNotifications(notifications);
          }

          // Update task escalation level
          const maxLevel = Math.max(...notifications.map(n => this.getEscalationLevel(n)));
          if (maxLevel > (task.EscalationLevel || 0)) {
            await this.updateTaskEscalationLevel(task.Id, maxLevel);
          }
        }
      } catch (error) {
        const errorMsg = `Error processing escalation for task ${task.Id}: ${error instanceof Error ? error.message : 'Unknown error'}`;
        errors.push(errorMsg);
        logger.warn('ScheduledTaskProcessor', errorMsg);
      }
    }

    return { escalationsGenerated, errors };
  }

  /**
   * Get escalation level from notification
   */
  private getEscalationLevel(notification: ITaskNotificationQueueItem): number {
    switch (notification.Priority) {
      case 'Urgent': return 3;
      case 'High': return 2;
      case 'Normal': return 1;
      default: return 1;
    }
  }

  /**
   * Update task escalation level
   */
  private async updateTaskEscalationLevel(taskId: number, level: number): Promise<void> {
    try {
      await this.sp.web.lists
        .getByTitle(this.tasksListTitle)
        .items.getById(taskId)
        .update({
          EscalationLevel: level,
          LastEscalationDate: new Date().toISOString()
        });
    } catch (error) {
      logger.warn('ScheduledTaskProcessor', `Failed to update escalation level for task ${taskId}`, error);
    }
  }

  // ============================================================================
  // SLA ENFORCEMENT
  // ============================================================================

  /**
   * Process SLA enforcement for all tasks
   */
  private async processSLAEnforcement(
    tasks: IJmlTaskAssignment[],
    config: IScheduledProcessorConfig
  ): Promise<{ breaches: number; warnings: number; errors: string[] }> {
    let breaches = 0;
    let warnings = 0;
    const errors: string[] = [];
    const now = new Date();

    for (const task of tasks) {
      try {
        // Skip tasks without SLA or blocked tasks
        if (!task.SLAHours || task.IsBlocked) {
          continue;
        }

        const slaStatus = this.calculateSLAStatus(task, now);

        if (slaStatus.status === 'Breached') {
          breaches++;
          if (!config.dryRun) {
            await this.handleSLABreach(task, slaStatus);
          }
        } else if (slaStatus.status === 'Warning') {
          warnings++;
          if (!config.dryRun) {
            await this.handleSLAWarning(task, slaStatus);
          }
        }
      } catch (error) {
        const errorMsg = `Error processing SLA for task ${task.Id}: ${error instanceof Error ? error.message : 'Unknown error'}`;
        errors.push(errorMsg);
        logger.warn('ScheduledTaskProcessor', errorMsg);
      }
    }

    return { breaches, warnings, errors };
  }

  /**
   * Calculate SLA status for a task
   */
  private calculateSLAStatus(task: IJmlTaskAssignment, now: Date): ITaskSLAStatus {
    const slaHours = task.SLAHours || 0;
    const startDate = task.AssignedDate || task.Created;
    const elapsedMs = now.getTime() - startDate.getTime();
    const elapsedHours = elapsedMs / (1000 * 60 * 60);
    const remainingHours = slaHours - elapsedHours;
    const percentUsed = (elapsedHours / slaHours) * 100;

    let status: 'Healthy' | 'Warning' | 'Breached' = 'Healthy';
    if (remainingHours <= 0) {
      status = 'Breached';
    } else if (percentUsed >= 80) {
      status = 'Warning';
    }

    return {
      taskId: task.Id,
      taskTitle: task.Title,
      status,
      slaHours,
      elapsedHours: Math.round(elapsedHours * 10) / 10,
      remainingHours: Math.round(remainingHours * 10) / 10,
      assigneeId: task.AssignedToId,
      assigneeEmail: task.AssignedTo?.EMail
    };
  }

  /**
   * Handle SLA breach - escalate and notify
   */
  private async handleSLABreach(task: IJmlTaskAssignment, slaStatus: ITaskSLAStatus): Promise<void> {
    try {
      // Check if we already logged this breach today
      const alreadyLogged = await this.checkSLALoggedToday(task.Id, 'Breach');
      if (alreadyLogged) {
        return;
      }

      // Log the breach
      await this.logSLAEvent(task, 'Breach', slaStatus);

      // Update task
      await this.sp.web.lists
        .getByTitle(this.tasksListTitle)
        .items.getById(task.Id)
        .update({
          SLAStatus: 'Breached',
          SLABreachedDate: new Date().toISOString(),
          Priority: task.Priority === 'Critical' ? 'Critical' : 'High' // Escalate priority
        });

      // Create notification for assignee and manager
      const recipients: number[] = [];
      if (task.AssignedToId) {
        recipients.push(task.AssignedToId);
      }

      const notification: ITaskNotificationQueueItem = {
        TaskAssignmentId: task.Id,
        NotificationType: TaskNotificationType.SLABreach,
        ScheduledFor: new Date(),
        Priority: 'Urgent',
        Recipients: recipients,
        Message: this.buildSLABreachMessage(task, slaStatus),
        IsProcessed: false
      };

      await this.queueNotifications([notification]);

      logger.info('ScheduledTaskProcessor', `SLA breach handled for task ${task.Id}`);
    } catch (error) {
      logger.error('ScheduledTaskProcessor', `Error handling SLA breach for task ${task.Id}`, error);
      throw error;
    }
  }

  /**
   * Handle SLA warning - notify before breach
   */
  private async handleSLAWarning(task: IJmlTaskAssignment, slaStatus: ITaskSLAStatus): Promise<void> {
    try {
      // Check if we already logged this warning today
      const alreadyLogged = await this.checkSLALoggedToday(task.Id, 'Warning');
      if (alreadyLogged) {
        return;
      }

      // Log the warning
      await this.logSLAEvent(task, 'Warning', slaStatus);

      // Update task
      await this.sp.web.lists
        .getByTitle(this.tasksListTitle)
        .items.getById(task.Id)
        .update({
          SLAStatus: 'Warning'
        });

      // Create notification for assignee
      const recipients: number[] = [];
      if (task.AssignedToId) {
        recipients.push(task.AssignedToId);
      }

      const notification: ITaskNotificationQueueItem = {
        TaskAssignmentId: task.Id,
        NotificationType: TaskNotificationType.SLAWarning,
        ScheduledFor: new Date(),
        Priority: 'High',
        Recipients: recipients,
        Message: this.buildSLAWarningMessage(task, slaStatus),
        IsProcessed: false
      };

      await this.queueNotifications([notification]);

      logger.info('ScheduledTaskProcessor', `SLA warning sent for task ${task.Id}`);
    } catch (error) {
      logger.error('ScheduledTaskProcessor', `Error handling SLA warning for task ${task.Id}`, error);
      throw error;
    }
  }

  /**
   * Build SLA breach notification message
   */
  private buildSLABreachMessage(task: IJmlTaskAssignment, slaStatus: ITaskSLAStatus): string {
    return `🔴 SLA BREACH: Task "${task.Title}" has exceeded its SLA.\n\n` +
      `SLA: ${slaStatus.slaHours} hours\n` +
      `Elapsed: ${slaStatus.elapsedHours} hours\n` +
      `Overdue by: ${Math.abs(slaStatus.remainingHours)} hours\n\n` +
      `Please complete this task immediately or escalate to your manager.`;
  }

  /**
   * Build SLA warning notification message
   */
  private buildSLAWarningMessage(task: IJmlTaskAssignment, slaStatus: ITaskSLAStatus): string {
    return `⚠️ SLA WARNING: Task "${task.Title}" is approaching its SLA deadline.\n\n` +
      `SLA: ${slaStatus.slaHours} hours\n` +
      `Remaining: ${slaStatus.remainingHours} hours\n` +
      `Elapsed: ${slaStatus.elapsedHours} hours\n\n` +
      `Please prioritize completing this task to avoid SLA breach.`;
  }

  /**
   * Check if SLA was already logged today
   */
  private async checkSLALoggedToday(taskId: number, eventType: string): Promise<boolean> {
    try {
      const today = new Date();
      today.setHours(0, 0, 0, 0);

      const items = await this.sp.web.lists
        .getByTitle(this.slaLogListTitle)
        .items.filter(
          `TaskAssignmentId eq ${taskId} and EventType eq '${eventType}' and EventDate ge datetime'${today.toISOString()}'`
        )
        .top(1)();

      return items.length > 0;
    } catch {
      // If list doesn't exist or query fails, assume not logged
      return false;
    }
  }

  /**
   * Log SLA event
   */
  private async logSLAEvent(task: IJmlTaskAssignment, eventType: string, slaStatus: ITaskSLAStatus): Promise<void> {
    try {
      await this.sp.web.lists
        .getByTitle(this.slaLogListTitle)
        .items.add({
          Title: `${eventType}: ${task.Title}`,
          TaskAssignmentId: task.Id,
          EventType: eventType,
          EventDate: new Date().toISOString(),
          SLAHours: slaStatus.slaHours,
          ElapsedHours: slaStatus.elapsedHours,
          RemainingHours: slaStatus.remainingHours,
          AssignedToId: task.AssignedToId
        });
    } catch (error) {
      logger.warn('ScheduledTaskProcessor', `Failed to log SLA event for task ${task.Id}`, error);
    }
  }

  // ============================================================================
  // NOTIFICATION QUEUE PROCESSING
  // ============================================================================

  /**
   * Queue notifications for later processing
   */
  private async queueNotifications(notifications: ITaskNotificationQueueItem[]): Promise<void> {
    for (const notification of notifications) {
      try {
        await this.sp.web.lists
          .getByTitle('JML_NotificationQueue')
          .items.add({
            Title: `Task ${notification.TaskAssignmentId}: ${notification.NotificationType}`,
            TaskAssignmentId: notification.TaskAssignmentId,
            NotificationType: notification.NotificationType,
            ScheduledFor: notification.ScheduledFor.toISOString(),
            Priority: notification.Priority,
            Recipients: JSON.stringify(notification.Recipients),
            Message: notification.Message,
            IsProcessed: false
          });
      } catch (error) {
        logger.warn('ScheduledTaskProcessor', 'Failed to queue notification', error);
      }
    }
  }

  /**
   * Process queued notifications - send emails and Teams messages
   */
  private async processNotificationQueue(): Promise<{
    sent: number;
    emailsSent: number;
    teamsMessagesSent: number;
    errors: string[];
  }> {
    let sent = 0;
    let emailsSent = 0;
    let teamsMessagesSent = 0;
    const errors: string[] = [];

    try {
      // Get unprocessed notifications
      const queueItems = await this.sp.web.lists
        .getByTitle('JML_NotificationQueue')
        .items.filter('IsProcessed eq false')
        .select('Id', 'TaskAssignmentId', 'NotificationType', 'Priority', 'Recipients', 'Message')
        .top(50)();

      for (const item of queueItems) {
        try {
          const recipients: number[] = JSON.parse(item.Recipients || '[]');

          // Get recipient emails
          const recipientEmails = await this.getRecipientEmails(recipients);

          if (recipientEmails.length > 0) {
            // Send email
            const emailSent = await this.sendNotificationEmail(
              recipientEmails,
              `JML Task Notification: ${item.NotificationType}`,
              item.Message,
              item.Priority
            );

            if (emailSent) {
              emailsSent++;
            }
          }

          // Create in-app notification
          for (const recipientId of recipients) {
            await this.createInAppNotification(recipientId, item);
          }

          sent += recipients.length;

          // Mark as processed
          await this.sp.web.lists
            .getByTitle('JML_NotificationQueue')
            .items.getById(item.Id)
            .update({
              IsProcessed: true,
              ProcessedDate: new Date().toISOString()
            });
        } catch (error) {
          const errorMsg = `Error processing notification ${item.Id}: ${error instanceof Error ? error.message : 'Unknown error'}`;
          errors.push(errorMsg);
          logger.warn('ScheduledTaskProcessor', errorMsg);
        }
      }
    } catch (error) {
      errors.push(`Error fetching notification queue: ${error instanceof Error ? error.message : 'Unknown error'}`);
    }

    return { sent, emailsSent, teamsMessagesSent, errors };
  }

  /**
   * Get email addresses for recipient IDs
   */
  private async getRecipientEmails(recipientIds: number[]): Promise<string[]> {
    const emails: string[] = [];

    for (const userId of recipientIds) {
      try {
        const user = await this.sp.web.siteUsers.getById(userId).select('Email')();
        if (user.Email) {
          emails.push(user.Email);
        }
      } catch {
        logger.warn('ScheduledTaskProcessor', `Could not get email for user ${userId}`);
      }
    }

    return emails;
  }

  /**
   * Send notification email via Graph API
   */
  private async sendNotificationEmail(
    to: string[],
    subject: string,
    body: string,
    priority: string
  ): Promise<boolean> {
    try {
      const graphClient: MSGraphClientV3 = await this.context.msGraphClientFactory.getClient('3');

      const htmlBody = this.buildEmailHtml(subject, body, priority);

      const emailMessage = {
        message: {
          subject,
          body: {
            contentType: 'HTML',
            content: htmlBody
          },
          toRecipients: to.map(email => ({
            emailAddress: { address: email }
          })),
          importance: priority === 'Urgent' ? 'high' : priority === 'High' ? 'high' : 'normal'
        },
        saveToSentItems: true
      };

      await graphClient.api('/me/sendMail').post(emailMessage);

      logger.info('ScheduledTaskProcessor', `Sent email to ${to.length} recipients`);
      return true;
    } catch (error) {
      logger.error('ScheduledTaskProcessor', 'Failed to send email', error);
      return false;
    }
  }

  /**
   * Build HTML email content
   */
  private buildEmailHtml(subject: string, body: string, priority: string): string {
    const priorityColor = priority === 'Urgent' ? '#d13438' : priority === 'High' ? '#ff8c00' : '#0078d4';
    const priorityBg = priority === 'Urgent' ? '#fde7e9' : priority === 'High' ? '#fff4ce' : '#e7f3ff';

    return `
<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <style>
    body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 0; padding: 0; background-color: #f3f2f1; }
    .container { max-width: 600px; margin: 0 auto; padding: 20px; }
    .card { background: #fff; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); overflow: hidden; }
    .header { background: ${priorityBg}; border-left: 4px solid ${priorityColor}; padding: 16px 20px; }
    .header h1 { margin: 0; font-size: 18px; color: ${priorityColor}; }
    .content { padding: 20px; }
    .content p { margin: 0 0 16px; color: #323130; line-height: 1.6; white-space: pre-line; }
    .button { display: inline-block; background: #0078d4; color: #fff; padding: 10px 24px; border-radius: 4px; text-decoration: none; font-weight: 600; }
    .footer { padding: 16px 20px; background: #faf9f8; text-align: center; font-size: 12px; color: #605e5c; }
  </style>
</head>
<body>
  <div class="container">
    <div class="card">
      <div class="header">
        <h1>${subject}</h1>
      </div>
      <div class="content">
        <p>${body}</p>
        <a href="${this.context.pageContext.web.absoluteUrl}/SitePages/MyTasks.aspx" class="button">View My Tasks</a>
      </div>
      <div class="footer">
        This is an automated notification from the JML System.
      </div>
    </div>
  </div>
</body>
</html>`;
  }

  /**
   * Create in-app notification
   */
  private async createInAppNotification(
    recipientId: number,
    queueItem: Record<string, unknown>
  ): Promise<void> {
    try {
      await this.sp.web.lists
        .getByTitle(this.notificationsListTitle)
        .items.add({
          Title: queueItem.NotificationType as string,
          Message: queueItem.Message as string,
          RecipientId: recipientId,
          Type: queueItem.NotificationType as string,
          Priority: queueItem.Priority as string,
          IsRead: false,
          RelatedItemType: 'Task',
          RelatedItemId: queueItem.TaskAssignmentId as number
        });
    } catch (error) {
      logger.warn('ScheduledTaskProcessor', `Failed to create in-app notification for user ${recipientId}`, error);
    }
  }

  // ============================================================================
  // LOGGING & UTILITIES
  // ============================================================================

  /**
   * Log processing run results
   */
  private async logProcessingRun(result: IScheduledProcessingResult): Promise<void> {
    try {
      await this.sp.web.lists
        .getByTitle('JML_ScheduledProcessingLog')
        .items.add({
          Title: `Processing Run - ${result.processedAt.toISOString()}`,
          ProcessedAt: result.processedAt.toISOString(),
          Duration: result.duration,
          TasksProcessed: result.tasksProcessed,
          EscalationsGenerated: result.escalationsGenerated,
          NotificationsSent: result.notificationsSent,
          EmailsSent: result.emailsSent,
          TeamsMessagesSent: result.teamsMessagesSent,
          SLABreaches: result.slaBreaches,
          SLAWarnings: result.slaWarnings,
          Success: result.success,
          Errors: result.errors.join('\n')
        });
    } catch {
      // Logging failure shouldn't break the process
      logger.warn('ScheduledTaskProcessor', 'Failed to log processing run');
    }
  }

  // ============================================================================
  // PUBLIC UTILITY METHODS
  // ============================================================================

  /**
   * Get SLA status for all tasks (for dashboard/reporting)
   */
  public async getSLADashboard(): Promise<{
    healthy: number;
    warning: number;
    breached: number;
    tasks: ITaskSLAStatus[];
  }> {
    const tasks = await this.getIncompleteTasks(1000);
    const now = new Date();
    const statuses: ITaskSLAStatus[] = [];

    let healthy = 0;
    let warning = 0;
    let breached = 0;

    for (const task of tasks) {
      if (task.SLAHours) {
        const status = this.calculateSLAStatus(task, now);
        statuses.push(status);

        switch (status.status) {
          case 'Healthy': healthy++; break;
          case 'Warning': warning++; break;
          case 'Breached': breached++; break;
        }
      }
    }

    return { healthy, warning, breached, tasks: statuses };
  }

  /**
   * Manually trigger escalation check for a specific task
   */
  public async checkTaskEscalation(taskId: number): Promise<ITaskNotificationQueueItem[]> {
    const item = await this.sp.web.lists
      .getByTitle(this.tasksListTitle)
      .items.getById(taskId)
      .select(
        'Id', 'Title', 'Status', 'Priority', 'DueDate', 'AssignedToId',
        'AssignedDate', 'ProcessIDId', 'TaskIDId', 'Modified', 'Created',
        'AssignedTo/Id', 'AssignedTo/Title', 'AssignedTo/EMail'
      )
      .expand('AssignedTo')();

    const task = this.mapToTaskAssignment(item);
    return this.taskNotificationService.evaluateTaskEscalation(task);
  }
}
