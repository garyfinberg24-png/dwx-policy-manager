// @ts-nocheck
/* eslint-disable */
/**
 * WorkflowNotificationService
 * Unified notification service for workflow-related events
 * Consolidates notification logic and provides consistent templates
 *
 * Supports two email sending modes:
 * 1. Direct (default): Uses Graph API /me/sendMail - requires user context
 * 2. Queue-based: Uses EmailQueueService - works without user context for background operations
 */

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { MSGraphClientV3 } from '@microsoft/sp-http';

import {
  IWorkflowInstance,
  IStepConfig,
  WorkflowInstanceStatus
} from '../../models/IWorkflow';
import { logger } from '../LoggingService';
import { EmailQueueService, EmailPriority } from '../EmailQueueService';

/**
 * Workflow notification event types
 */
export enum WorkflowNotificationEvent {
  // Lifecycle events
  WorkflowStarted = 'WorkflowStarted',
  WorkflowCompleted = 'WorkflowCompleted',
  WorkflowFailed = 'WorkflowFailed',
  WorkflowCancelled = 'WorkflowCancelled',

  // Step events
  StepStarted = 'StepStarted',
  StepCompleted = 'StepCompleted',
  StepFailed = 'StepFailed',

  // Task events
  TaskAssigned = 'TaskAssigned',
  TaskReminder = 'TaskReminder',
  TaskOverdue = 'TaskOverdue',
  TaskCompleted = 'TaskCompleted',
  TaskEscalated = 'TaskEscalated',

  // Approval events
  ApprovalRequested = 'ApprovalRequested',
  ApprovalReminder = 'ApprovalReminder',
  ApprovalCompleted = 'ApprovalCompleted',
  ApprovalRejected = 'ApprovalRejected',
  ApprovalDelegated = 'ApprovalDelegated',
  ApprovalEscalated = 'ApprovalEscalated',

  // SLA events
  SLAWarning = 'SLAWarning',
  SLABreach = 'SLABreach',

  // General
  Custom = 'Custom'
}

/**
 * Notification channel types
 */
export enum NotificationChannel {
  InApp = 'InApp',
  Email = 'Email',
  Teams = 'Teams',
  All = 'All'
}

/**
 * Notification priority levels
 */
export enum NotificationPriority {
  Low = 'Low',
  Normal = 'Normal',
  High = 'High',
  Urgent = 'Urgent'
}

/**
 * Notification request structure
 */
export interface IWorkflowNotification {
  event: WorkflowNotificationEvent;
  recipientIds: number[];
  recipientEmails?: string[];
  workflowInstance: IWorkflowInstance;
  stepConfig?: IStepConfig;
  channel?: NotificationChannel;
  priority?: NotificationPriority;
  customSubject?: string;
  customMessage?: string;
  additionalData?: Record<string, unknown>;
  teamsConfig?: {
    teamId?: string;
    channelId?: string;
    mentionUsers?: boolean;
  };
}

/**
 * Notification result
 */
export interface INotificationResult {
  success: boolean;
  notificationIds?: number[];
  emailsSent?: number;
  emailsQueued?: number;
  teamsMessagesSent?: number;
  errors?: string[];
}

/**
 * Email sending mode
 */
export enum EmailSendMode {
  /** Send email directly via Graph API - requires user context */
  Direct = 'Direct',
  /** Queue email for background processing - works without user context */
  Queue = 'Queue',
  /** Automatically detect: use Direct if context available, Queue otherwise */
  Auto = 'Auto'
}

/**
 * Service configuration options
 */
export interface IWorkflowNotificationServiceConfig {
  /** Email sending mode (default: Auto) */
  emailMode?: EmailSendMode;
  /** Force queue mode even when context is available */
  forceQueueMode?: boolean;
}

export class WorkflowNotificationService {
  private sp: SPFI;
  private context: WebPartContext | undefined;
  private notificationsListTitle = 'JML_Notifications';
  private siteUrl: string;
  private emailMode: EmailSendMode;
  private forceQueueMode: boolean;
  private emailQueueService: EmailQueueService | undefined;

  /**
   * Create notification service with WebPartContext (supports direct email)
   */
  constructor(sp: SPFI, context: WebPartContext, config?: IWorkflowNotificationServiceConfig);
  /**
   * Create notification service without context (queue-only mode)
   */
  constructor(sp: SPFI, siteUrl: string, config?: IWorkflowNotificationServiceConfig);
  constructor(
    sp: SPFI,
    contextOrSiteUrl: WebPartContext | string,
    config?: IWorkflowNotificationServiceConfig
  ) {
    this.sp = sp;
    this.emailMode = config?.emailMode || EmailSendMode.Auto;
    this.forceQueueMode = config?.forceQueueMode || false;

    if (typeof contextOrSiteUrl === 'string') {
      // No context - queue-only mode
      this.context = undefined;
      this.siteUrl = contextOrSiteUrl;
      this.emailMode = EmailSendMode.Queue; // Force queue mode
    } else {
      // Has context - can use direct mode
      this.context = contextOrSiteUrl;
      this.siteUrl = contextOrSiteUrl.pageContext.web.absoluteUrl;
    }

    // Initialize email queue service for queue-based sending
    this.emailQueueService = new EmailQueueService(sp);
  }

  /**
   * Check if direct email sending is available
   */
  public canSendDirectEmail(): boolean {
    return this.context !== undefined && !this.forceQueueMode;
  }

  /**
   * Get the effective email mode based on configuration and context
   */
  private getEffectiveEmailMode(): EmailSendMode {
    if (this.forceQueueMode) {
      return EmailSendMode.Queue;
    }

    if (this.emailMode === EmailSendMode.Auto) {
      return this.context ? EmailSendMode.Direct : EmailSendMode.Queue;
    }

    // If Direct mode requested but no context, fall back to Queue
    if (this.emailMode === EmailSendMode.Direct && !this.context) {
      logger.warn('WorkflowNotificationService', 'Direct email mode requested but no context available, falling back to Queue');
      return EmailSendMode.Queue;
    }

    return this.emailMode;
  }

  // ============================================================================
  // PUBLIC METHODS - LIFECYCLE NOTIFICATIONS
  // ============================================================================

  /**
   * Send workflow started notification
   */
  public async notifyWorkflowStarted(
    workflowInstance: IWorkflowInstance,
    recipientIds: number[],
    recipientEmails?: string[]
  ): Promise<INotificationResult> {
    return this.sendNotification({
      event: WorkflowNotificationEvent.WorkflowStarted,
      recipientIds,
      recipientEmails,
      workflowInstance,
      channel: NotificationChannel.All,
      priority: NotificationPriority.Normal
    });
  }

  /**
   * Send workflow completed notification
   */
  public async notifyWorkflowCompleted(
    workflowInstance: IWorkflowInstance,
    recipientIds: number[],
    recipientEmails?: string[]
  ): Promise<INotificationResult> {
    return this.sendNotification({
      event: WorkflowNotificationEvent.WorkflowCompleted,
      recipientIds,
      recipientEmails,
      workflowInstance,
      channel: NotificationChannel.All,
      priority: NotificationPriority.Normal
    });
  }

  /**
   * Send workflow failed notification
   */
  public async notifyWorkflowFailed(
    workflowInstance: IWorkflowInstance,
    recipientIds: number[],
    errorMessage: string,
    recipientEmails?: string[]
  ): Promise<INotificationResult> {
    return this.sendNotification({
      event: WorkflowNotificationEvent.WorkflowFailed,
      recipientIds,
      recipientEmails,
      workflowInstance,
      channel: NotificationChannel.All,
      priority: NotificationPriority.High,
      additionalData: { errorMessage }
    });
  }

  // ============================================================================
  // PUBLIC METHODS - TASK NOTIFICATIONS
  // ============================================================================

  /**
   * Send task assigned notification
   */
  public async notifyTaskAssigned(
    workflowInstance: IWorkflowInstance,
    taskId: number,
    taskTitle: string,
    assigneeId: number,
    assigneeEmail?: string,
    dueDate?: Date
  ): Promise<INotificationResult> {
    return this.sendNotification({
      event: WorkflowNotificationEvent.TaskAssigned,
      recipientIds: [assigneeId],
      recipientEmails: assigneeEmail ? [assigneeEmail] : undefined,
      workflowInstance,
      channel: NotificationChannel.All,
      priority: NotificationPriority.Normal,
      additionalData: { taskId, taskTitle, dueDate: dueDate?.toISOString() }
    });
  }

  /**
   * Send task overdue notification
   */
  public async notifyTaskOverdue(
    workflowInstance: IWorkflowInstance,
    taskId: number,
    taskTitle: string,
    recipientIds: number[],
    hoursOverdue: number,
    recipientEmails?: string[]
  ): Promise<INotificationResult> {
    return this.sendNotification({
      event: WorkflowNotificationEvent.TaskOverdue,
      recipientIds,
      recipientEmails,
      workflowInstance,
      channel: NotificationChannel.All,
      priority: NotificationPriority.High,
      additionalData: { taskId, taskTitle, hoursOverdue }
    });
  }

  /**
   * Send task escalation notification
   */
  public async notifyTaskEscalated(
    workflowInstance: IWorkflowInstance,
    taskId: number,
    taskTitle: string,
    escalatedToIds: number[],
    escalationLevel: number,
    reason: string,
    recipientEmails?: string[]
  ): Promise<INotificationResult> {
    return this.sendNotification({
      event: WorkflowNotificationEvent.TaskEscalated,
      recipientIds: escalatedToIds,
      recipientEmails,
      workflowInstance,
      channel: NotificationChannel.All,
      priority: NotificationPriority.Urgent,
      additionalData: { taskId, taskTitle, escalationLevel, reason }
    });
  }

  // ============================================================================
  // PUBLIC METHODS - APPROVAL NOTIFICATIONS
  // ============================================================================

  /**
   * Send approval requested notification
   */
  public async notifyApprovalRequested(
    workflowInstance: IWorkflowInstance,
    approvalId: number,
    approvalTitle: string,
    approverId: number,
    approverEmail?: string,
    dueDate?: Date
  ): Promise<INotificationResult> {
    return this.sendNotification({
      event: WorkflowNotificationEvent.ApprovalRequested,
      recipientIds: [approverId],
      recipientEmails: approverEmail ? [approverEmail] : undefined,
      workflowInstance,
      channel: NotificationChannel.All,
      priority: NotificationPriority.High,
      additionalData: { approvalId, approvalTitle, dueDate: dueDate?.toISOString() }
    });
  }

  /**
   * Send approval completed notification
   */
  public async notifyApprovalCompleted(
    workflowInstance: IWorkflowInstance,
    approvalId: number,
    approvalTitle: string,
    approved: boolean,
    comments: string | undefined,
    recipientIds: number[],
    recipientEmails?: string[]
  ): Promise<INotificationResult> {
    return this.sendNotification({
      event: approved ? WorkflowNotificationEvent.ApprovalCompleted : WorkflowNotificationEvent.ApprovalRejected,
      recipientIds,
      recipientEmails,
      workflowInstance,
      channel: NotificationChannel.All,
      priority: approved ? NotificationPriority.Normal : NotificationPriority.High,
      additionalData: { approvalId, approvalTitle, approved, comments }
    });
  }

  // ============================================================================
  // PUBLIC METHODS - SLA NOTIFICATIONS
  // ============================================================================

  /**
   * Send SLA warning notification
   */
  public async notifySLAWarning(
    workflowInstance: IWorkflowInstance,
    stepName: string,
    hoursRemaining: number,
    recipientIds: number[],
    recipientEmails?: string[]
  ): Promise<INotificationResult> {
    return this.sendNotification({
      event: WorkflowNotificationEvent.SLAWarning,
      recipientIds,
      recipientEmails,
      workflowInstance,
      channel: NotificationChannel.All,
      priority: NotificationPriority.High,
      additionalData: { stepName, hoursRemaining }
    });
  }

  /**
   * Send SLA breach notification
   */
  public async notifySLABreach(
    workflowInstance: IWorkflowInstance,
    stepName: string,
    hoursOverdue: number,
    recipientIds: number[],
    recipientEmails?: string[]
  ): Promise<INotificationResult> {
    return this.sendNotification({
      event: WorkflowNotificationEvent.SLABreach,
      recipientIds,
      recipientEmails,
      workflowInstance,
      channel: NotificationChannel.All,
      priority: NotificationPriority.Urgent,
      additionalData: { stepName, hoursOverdue }
    });
  }

  // ============================================================================
  // CORE NOTIFICATION METHOD
  // ============================================================================

  /**
   * Send notification through specified channels
   */
  public async sendNotification(notification: IWorkflowNotification): Promise<INotificationResult> {
    const result: INotificationResult = {
      success: true,
      notificationIds: [],
      emailsSent: 0,
      emailsQueued: 0,
      teamsMessagesSent: 0,
      errors: []
    };

    const channel = notification.channel || NotificationChannel.InApp;
    const { subject, body, htmlBody } = this.buildNotificationContent(notification);

    try {
      // Send In-App notifications
      if (channel === NotificationChannel.InApp || channel === NotificationChannel.All) {
        const inAppResult = await this.sendInAppNotifications(
          notification.recipientIds,
          subject,
          body,
          notification
        );
        result.notificationIds = inAppResult.notificationIds;
        if (inAppResult.errors.length > 0) {
          result.errors?.push(...inAppResult.errors);
        }
      }

      // Send Email notifications (supports both direct and queue modes)
      if (channel === NotificationChannel.Email || channel === NotificationChannel.All) {
        if (notification.recipientEmails && notification.recipientEmails.length > 0) {
          const emailResult = await this.sendEmailNotifications(
            notification.recipientEmails,
            subject,
            htmlBody,
            notification
          );
          result.emailsSent = emailResult.sent;
          result.emailsQueued = emailResult.queued;
          if (emailResult.errors.length > 0) {
            result.errors?.push(...emailResult.errors);
          }
        }
      }

      // Send Teams notifications
      if (channel === NotificationChannel.Teams || channel === NotificationChannel.All) {
        if (notification.teamsConfig) {
          const teamsResult = await this.sendTeamsNotification(
            notification.teamsConfig,
            subject,
            body,
            notification
          );
          result.teamsMessagesSent = teamsResult.sent;
          if (teamsResult.errors.length > 0) {
            result.errors?.push(...teamsResult.errors);
          }
        }
      }

      // Success if no errors, or if emails were at least queued successfully
      const hasEmailSuccess = (result.emailsSent || 0) > 0 || (result.emailsQueued || 0) > 0;
      const hasOnlyQueueErrors = result.errors?.every(e => e.includes('Failed to send email directly')) || false;
      result.success = (result.errors?.length || 0) === 0 || (hasEmailSuccess && hasOnlyQueueErrors);

    } catch (error) {
      result.success = false;
      result.errors?.push(error instanceof Error ? error.message : 'Unknown error sending notification');
      logger.error('WorkflowNotificationService', 'Error sending notification', error);
    }

    return result;
  }

  // ============================================================================
  // PRIVATE METHODS - CHANNEL HANDLERS
  // ============================================================================

  /**
   * Send in-app notifications
   */
  private async sendInAppNotifications(
    recipientIds: number[],
    title: string,
    message: string,
    notification: IWorkflowNotification
  ): Promise<{ notificationIds: number[]; errors: string[] }> {
    const notificationIds: number[] = [];
    const errors: string[] = [];

    for (const recipientId of recipientIds) {
      try {
        const notificationData = {
          Title: title,
          Message: message,
          RecipientId: recipientId,
          Type: this.mapEventToNotificationType(notification.event),
          Priority: notification.priority || 'Normal',
          IsRead: false,
          RelatedItemType: 'WorkflowInstance',
          RelatedItemId: notification.workflowInstance.Id,
          WorkflowInstanceId: notification.workflowInstance.Id,
          ProcessId: notification.workflowInstance.ProcessId,
          EventType: notification.event,
          ActionUrl: this.buildActionUrl(notification)
        };

        const result = await this.sp.web.lists
          .getByTitle(this.notificationsListTitle)
          .items.add(notificationData);

        notificationIds.push(result.data.Id);
      } catch (error) {
        errors.push(`Failed to create notification for user ${recipientId}: ${error instanceof Error ? error.message : 'Unknown error'}`);
        logger.warn('WorkflowNotificationService', `Failed to create in-app notification for user ${recipientId}`, error);
      }
    }

    return { notificationIds, errors };
  }

  /**
   * Send email notifications via Graph API (direct) or EmailQueueService (queue)
   * Automatically selects the appropriate method based on configuration and context
   */
  private async sendEmailNotifications(
    recipientEmails: string[],
    subject: string,
    htmlBody: string,
    notification?: IWorkflowNotification
  ): Promise<{ sent: number; queued: number; errors: string[] }> {
    let sent = 0;
    let queued = 0;
    const errors: string[] = [];

    const effectiveMode = this.getEffectiveEmailMode();

    if (effectiveMode === EmailSendMode.Direct) {
      // Direct mode: Send via Graph API
      try {
        if (!this.context) {
          throw new Error('No context available for direct email sending');
        }

        const emailMessage = {
          message: {
            subject,
            body: {
              contentType: 'HTML',
              content: htmlBody
            },
            toRecipients: recipientEmails.map(email => ({
              emailAddress: { address: email }
            }))
          },
          saveToSentItems: true
        };

        const graphClient: MSGraphClientV3 = await this.context.msGraphClientFactory.getClient('3');
        await graphClient.api('/me/sendMail').post(emailMessage);

        sent = recipientEmails.length;
        logger.info('WorkflowNotificationService', `Sent email directly to ${recipientEmails.length} recipients`);
      } catch (error) {
        errors.push(`Failed to send email directly: ${error instanceof Error ? error.message : 'Unknown error'}`);
        logger.error('WorkflowNotificationService', 'Failed to send email notification via Graph', error);

        // Fall back to queue on direct failure
        if (this.emailQueueService) {
          logger.info('WorkflowNotificationService', 'Falling back to queue-based email sending');
          const queueResult = await this.queueEmailNotification(recipientEmails, subject, htmlBody, notification);
          queued = queueResult.queued;
          if (queueResult.errors.length > 0) {
            errors.push(...queueResult.errors);
          }
        }
      }
    } else {
      // Queue mode: Use EmailQueueService
      const queueResult = await this.queueEmailNotification(recipientEmails, subject, htmlBody, notification);
      queued = queueResult.queued;
      if (queueResult.errors.length > 0) {
        errors.push(...queueResult.errors);
      }
    }

    return { sent, queued, errors };
  }

  /**
   * Queue email notification via EmailQueueService
   */
  private async queueEmailNotification(
    recipientEmails: string[],
    subject: string,
    htmlBody: string,
    notification?: IWorkflowNotification
  ): Promise<{ queued: number; errors: string[] }> {
    let queued = 0;
    const errors: string[] = [];

    if (!this.emailQueueService) {
      errors.push('EmailQueueService not initialized');
      return { queued, errors };
    }

    try {
      // Map notification priority to email priority
      let priority = EmailPriority.Normal;
      if (notification?.priority) {
        switch (notification.priority) {
          case NotificationPriority.Low:
            priority = EmailPriority.Low;
            break;
          case NotificationPriority.High:
            priority = EmailPriority.High;
            break;
          case NotificationPriority.Urgent:
            priority = EmailPriority.Urgent;
            break;
          default:
            priority = EmailPriority.Normal;
        }
      }

      const result = await this.emailQueueService.queueEmail({
        to: recipientEmails,
        subject,
        htmlBody,
        priority,
        processId: notification?.workflowInstance?.ProcessId,
        workflowInstanceId: notification?.workflowInstance?.Id,
        notificationType: notification?.event
      });

      if (result.success) {
        queued = recipientEmails.length;
        logger.info('WorkflowNotificationService', `Queued email for ${recipientEmails.length} recipients (Queue ID: ${result.queueItemId})`);
      } else {
        errors.push(result.error || 'Failed to queue email');
      }
    } catch (error) {
      errors.push(`Failed to queue email: ${error instanceof Error ? error.message : 'Unknown error'}`);
      logger.error('WorkflowNotificationService', 'Failed to queue email notification', error);
    }

    return { queued, errors };
  }

  /**
   * Send Teams notification
   */
  private async sendTeamsNotification(
    teamsConfig: { teamId?: string; channelId?: string; mentionUsers?: boolean },
    subject: string,
    body: string,
    notification: IWorkflowNotification
  ): Promise<{ sent: number; errors: string[] }> {
    let sent = 0;
    const errors: string[] = [];

    if (!teamsConfig.teamId || !teamsConfig.channelId) {
      return { sent, errors };
    }

    try {
      const graphClient: MSGraphClientV3 = await this.context.msGraphClientFactory.getClient('3');

      const teamsMessage = {
        body: {
          contentType: 'html',
          content: `<strong>${subject}</strong><br/><br/>${body}`
        }
      };

      await graphClient
        .api(`/teams/${teamsConfig.teamId}/channels/${teamsConfig.channelId}/messages`)
        .post(teamsMessage);

      sent = 1;
      logger.info('WorkflowNotificationService', `Sent Teams message to channel ${teamsConfig.channelId}`);
    } catch (error) {
      errors.push(`Failed to send Teams message: ${error instanceof Error ? error.message : 'Unknown error'}`);
      logger.error('WorkflowNotificationService', 'Failed to send Teams notification', error);
    }

    return { sent, errors };
  }

  // ============================================================================
  // PRIVATE METHODS - CONTENT BUILDING
  // ============================================================================

  /**
   * Build notification content based on event type
   */
  private buildNotificationContent(notification: IWorkflowNotification): {
    subject: string;
    body: string;
    htmlBody: string;
  } {
    const instance = notification.workflowInstance;
    const processType = instance.ProcessType || 'Process';
    const data = notification.additionalData || {};

    let subject = '';
    let body = '';
    let htmlBody = '';

    switch (notification.event) {
      // Lifecycle events
      case WorkflowNotificationEvent.WorkflowStarted:
        subject = `Workflow Started: ${processType} #${instance.ProcessId}`;
        body = `A new ${processType} workflow has been initiated and is now in progress.`;
        htmlBody = this.buildHtmlTemplate(subject, body, 'info', instance);
        break;

      case WorkflowNotificationEvent.WorkflowCompleted:
        subject = `Workflow Completed: ${processType} #${instance.ProcessId}`;
        body = `The ${processType} workflow has been successfully completed.`;
        htmlBody = this.buildHtmlTemplate(subject, body, 'success', instance);
        break;

      case WorkflowNotificationEvent.WorkflowFailed:
        subject = `Workflow Failed: ${processType} #${instance.ProcessId}`;
        body = `The ${processType} workflow has failed. Error: ${data.errorMessage || 'Unknown error'}`;
        htmlBody = this.buildHtmlTemplate(subject, body, 'error', instance);
        break;

      case WorkflowNotificationEvent.WorkflowCancelled:
        subject = `Workflow Cancelled: ${processType} #${instance.ProcessId}`;
        body = `The ${processType} workflow has been cancelled.`;
        htmlBody = this.buildHtmlTemplate(subject, body, 'warning', instance);
        break;

      // Task events
      case WorkflowNotificationEvent.TaskAssigned:
        subject = `Task Assigned: ${data.taskTitle}`;
        body = `You have been assigned a new task: ${data.taskTitle}${data.dueDate ? `. Due: ${new Date(data.dueDate as string).toLocaleDateString()}` : ''}`;
        htmlBody = this.buildTaskHtmlTemplate(subject, body, data, instance);
        break;

      case WorkflowNotificationEvent.TaskOverdue:
        subject = `Task Overdue: ${data.taskTitle}`;
        body = `Task "${data.taskTitle}" is ${data.hoursOverdue} hours overdue and requires immediate attention.`;
        htmlBody = this.buildHtmlTemplate(subject, body, 'error', instance);
        break;

      case WorkflowNotificationEvent.TaskEscalated:
        subject = `Task Escalated (Level ${data.escalationLevel}): ${data.taskTitle}`;
        body = `Task "${data.taskTitle}" has been escalated to you. Reason: ${data.reason}`;
        htmlBody = this.buildHtmlTemplate(subject, body, 'warning', instance);
        break;

      case WorkflowNotificationEvent.TaskCompleted:
        subject = `Task Completed: ${data.taskTitle}`;
        body = `Task "${data.taskTitle}" has been marked as complete.`;
        htmlBody = this.buildHtmlTemplate(subject, body, 'success', instance);
        break;

      // Approval events
      case WorkflowNotificationEvent.ApprovalRequested:
        subject = `Approval Required: ${data.approvalTitle}`;
        body = `Your approval is required for: ${data.approvalTitle}${data.dueDate ? `. Please respond by: ${new Date(data.dueDate as string).toLocaleDateString()}` : ''}`;
        htmlBody = this.buildApprovalHtmlTemplate(subject, body, data, instance, true);
        break;

      case WorkflowNotificationEvent.ApprovalCompleted:
        subject = `Approval Granted: ${data.approvalTitle}`;
        body = `The approval request "${data.approvalTitle}" has been approved.${data.comments ? ` Comments: ${data.comments}` : ''}`;
        htmlBody = this.buildHtmlTemplate(subject, body, 'success', instance);
        break;

      case WorkflowNotificationEvent.ApprovalRejected:
        subject = `Approval Rejected: ${data.approvalTitle}`;
        body = `The approval request "${data.approvalTitle}" has been rejected.${data.comments ? ` Reason: ${data.comments}` : ''}`;
        htmlBody = this.buildHtmlTemplate(subject, body, 'error', instance);
        break;

      case WorkflowNotificationEvent.ApprovalEscalated:
        subject = `Approval Escalated: ${data.approvalTitle}`;
        body = `The approval request "${data.approvalTitle}" has been escalated to you for review.`;
        htmlBody = this.buildApprovalHtmlTemplate(subject, body, data, instance, true);
        break;

      // SLA events
      case WorkflowNotificationEvent.SLAWarning:
        subject = `SLA Warning: ${data.stepName}`;
        body = `The step "${data.stepName}" has ${data.hoursRemaining} hours remaining before SLA breach.`;
        htmlBody = this.buildHtmlTemplate(subject, body, 'warning', instance);
        break;

      case WorkflowNotificationEvent.SLABreach:
        subject = `SLA BREACH: ${data.stepName}`;
        body = `The step "${data.stepName}" has breached its SLA. It is ${data.hoursOverdue} hours overdue.`;
        htmlBody = this.buildHtmlTemplate(subject, body, 'error', instance);
        break;

      // Custom
      case WorkflowNotificationEvent.Custom:
        subject = notification.customSubject || 'Workflow Notification';
        body = notification.customMessage || 'You have a new workflow notification.';
        htmlBody = this.buildHtmlTemplate(subject, body, 'info', instance);
        break;

      default:
        subject = 'Workflow Notification';
        body = 'You have a new notification from the workflow system.';
        htmlBody = this.buildHtmlTemplate(subject, body, 'info', instance);
    }

    return { subject, body, htmlBody };
  }

  /**
   * Build standard HTML email template
   */
  private buildHtmlTemplate(
    subject: string,
    body: string,
    type: 'info' | 'success' | 'warning' | 'error',
    instance: IWorkflowInstance
  ): string {
    const colors = {
      info: { bg: '#e7f3ff', border: '#0078d4', text: '#004578' },
      success: { bg: '#dff6dd', border: '#107c10', text: '#0b5c0b' },
      warning: { bg: '#fff4ce', border: '#ff8c00', text: '#8a6914' },
      error: { bg: '#fde7e9', border: '#d13438', text: '#a80000' }
    };
    const color = colors[type];

    const processUrl = `${this.siteUrl}/SitePages/ProcessDetails.aspx?processId=${instance.ProcessId}`;

    return `
<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <style>
    body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 0; padding: 0; background-color: #f3f2f1; }
    .container { max-width: 600px; margin: 0 auto; padding: 20px; }
    .card { background: #fff; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); overflow: hidden; }
    .header { background: ${color.bg}; border-left: 4px solid ${color.border}; padding: 16px 20px; }
    .header h1 { margin: 0; font-size: 18px; color: ${color.text}; }
    .content { padding: 20px; }
    .content p { margin: 0 0 16px; color: #323130; line-height: 1.5; }
    .details { background: #faf9f8; border-radius: 4px; padding: 12px 16px; margin: 16px 0; }
    .details-row { display: flex; justify-content: space-between; padding: 4px 0; }
    .details-label { color: #605e5c; font-size: 13px; }
    .details-value { color: #323130; font-size: 13px; font-weight: 600; }
    .button { display: inline-block; background: #0078d4; color: #fff; padding: 10px 24px; border-radius: 4px; text-decoration: none; font-weight: 600; margin-top: 16px; }
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
        <div class="details">
          <div class="details-row">
            <span class="details-label">Process Type</span>
            <span class="details-value">${instance.ProcessType || 'N/A'}</span>
          </div>
          <div class="details-row">
            <span class="details-label">Process ID</span>
            <span class="details-value">#${instance.ProcessId}</span>
          </div>
          <div class="details-row">
            <span class="details-label">Status</span>
            <span class="details-value">${instance.Status}</span>
          </div>
        </div>
        <a href="${processUrl}" class="button">View Process Details</a>
      </div>
      <div class="footer">
        This is an automated notification from the JML Workflow System.
      </div>
    </div>
  </div>
</body>
</html>`;
  }

  /**
   * Build task-specific HTML template
   */
  private buildTaskHtmlTemplate(
    subject: string,
    body: string,
    data: Record<string, unknown>,
    instance: IWorkflowInstance
  ): string {
    const taskUrl = `${this.siteUrl}/SitePages/MyTasks.aspx`;

    return `
<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <style>
    body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 0; padding: 0; background-color: #f3f2f1; }
    .container { max-width: 600px; margin: 0 auto; padding: 20px; }
    .card { background: #fff; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); overflow: hidden; }
    .header { background: #e7f3ff; border-left: 4px solid #0078d4; padding: 16px 20px; }
    .header h1 { margin: 0; font-size: 18px; color: #004578; }
    .content { padding: 20px; }
    .content p { margin: 0 0 16px; color: #323130; line-height: 1.5; }
    .task-card { background: #faf9f8; border: 1px solid #edebe9; border-radius: 4px; padding: 16px; margin: 16px 0; }
    .task-title { font-size: 16px; font-weight: 600; color: #323130; margin-bottom: 8px; }
    .task-meta { display: flex; gap: 16px; font-size: 13px; color: #605e5c; }
    .button { display: inline-block; background: #0078d4; color: #fff; padding: 10px 24px; border-radius: 4px; text-decoration: none; font-weight: 600; margin-top: 16px; }
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
        <div class="task-card">
          <div class="task-title">${data.taskTitle}</div>
          <div class="task-meta">
            <span>Process: ${instance.ProcessType} #${instance.ProcessId}</span>
            ${data.dueDate ? `<span>Due: ${new Date(data.dueDate as string).toLocaleDateString()}</span>` : ''}
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

  /**
   * Build approval-specific HTML template
   */
  private buildApprovalHtmlTemplate(
    subject: string,
    body: string,
    data: Record<string, unknown>,
    instance: IWorkflowInstance,
    showActions: boolean
  ): string {
    const approvalUrl = `${this.siteUrl}/SitePages/ApprovalCenter.aspx`;

    return `
<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <style>
    body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 0; padding: 0; background-color: #f3f2f1; }
    .container { max-width: 600px; margin: 0 auto; padding: 20px; }
    .card { background: #fff; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); overflow: hidden; }
    .header { background: #fff4ce; border-left: 4px solid #ff8c00; padding: 16px 20px; }
    .header h1 { margin: 0; font-size: 18px; color: #8a6914; }
    .content { padding: 20px; }
    .content p { margin: 0 0 16px; color: #323130; line-height: 1.5; }
    .approval-card { background: #faf9f8; border: 1px solid #edebe9; border-radius: 4px; padding: 16px; margin: 16px 0; }
    .approval-title { font-size: 16px; font-weight: 600; color: #323130; margin-bottom: 8px; }
    .approval-meta { font-size: 13px; color: #605e5c; margin-bottom: 12px; }
    .actions { display: flex; gap: 8px; margin-top: 16px; }
    .btn-approve { display: inline-block; background: #107c10; color: #fff; padding: 10px 24px; border-radius: 4px; text-decoration: none; font-weight: 600; }
    .btn-reject { display: inline-block; background: #d13438; color: #fff; padding: 10px 24px; border-radius: 4px; text-decoration: none; font-weight: 600; }
    .btn-view { display: inline-block; background: #0078d4; color: #fff; padding: 10px 24px; border-radius: 4px; text-decoration: none; font-weight: 600; }
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
        <div class="approval-card">
          <div class="approval-title">${data.approvalTitle}</div>
          <div class="approval-meta">
            Process: ${instance.ProcessType} #${instance.ProcessId}
            ${data.dueDate ? `<br>Respond by: ${new Date(data.dueDate as string).toLocaleDateString()}` : ''}
          </div>
        </div>
        ${showActions ? `
        <div class="actions">
          <a href="${approvalUrl}" class="btn-view">Review & Respond</a>
        </div>
        ` : `<a href="${approvalUrl}" class="btn-view">View Approval Center</a>`}
      </div>
      <div class="footer">
        This is an automated notification from the JML Workflow System.
      </div>
    </div>
  </div>
</body>
</html>`;
  }

  // ============================================================================
  // PRIVATE METHODS - UTILITIES
  // ============================================================================

  /**
   * Map workflow event to notification type for in-app display
   */
  private mapEventToNotificationType(event: WorkflowNotificationEvent): string {
    switch (event) {
      case WorkflowNotificationEvent.TaskAssigned:
      case WorkflowNotificationEvent.TaskReminder:
      case WorkflowNotificationEvent.TaskCompleted:
        return 'Task';
      case WorkflowNotificationEvent.TaskOverdue:
      case WorkflowNotificationEvent.TaskEscalated:
        return 'Escalation';
      case WorkflowNotificationEvent.ApprovalRequested:
      case WorkflowNotificationEvent.ApprovalReminder:
      case WorkflowNotificationEvent.ApprovalCompleted:
      case WorkflowNotificationEvent.ApprovalRejected:
      case WorkflowNotificationEvent.ApprovalDelegated:
      case WorkflowNotificationEvent.ApprovalEscalated:
        return 'Approval';
      case WorkflowNotificationEvent.SLAWarning:
      case WorkflowNotificationEvent.SLABreach:
        return 'SLA';
      case WorkflowNotificationEvent.WorkflowFailed:
        return 'Error';
      default:
        return 'Info';
    }
  }

  /**
   * Build action URL based on notification type
   */
  private buildActionUrl(notification: IWorkflowNotification): string {
    const data = notification.additionalData || {};

    switch (notification.event) {
      case WorkflowNotificationEvent.TaskAssigned:
      case WorkflowNotificationEvent.TaskOverdue:
      case WorkflowNotificationEvent.TaskEscalated:
        return data.taskId ? `${this.siteUrl}/SitePages/MyTasks.aspx?taskId=${data.taskId}` : `${this.siteUrl}/SitePages/MyTasks.aspx`;

      case WorkflowNotificationEvent.ApprovalRequested:
      case WorkflowNotificationEvent.ApprovalEscalated:
        return data.approvalId ? `${this.siteUrl}/SitePages/ApprovalCenter.aspx?approvalId=${data.approvalId}` : `${this.siteUrl}/SitePages/ApprovalCenter.aspx`;

      default:
        return `${this.siteUrl}/SitePages/ProcessDetails.aspx?processId=${notification.workflowInstance.ProcessId}`;
    }
  }
}
