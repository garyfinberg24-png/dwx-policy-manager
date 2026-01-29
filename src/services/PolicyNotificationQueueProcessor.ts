// @ts-nocheck
/**
 * PolicyNotificationQueueProcessor
 * Processes policy-related notifications from PM_NotificationQueue
 *
 * Handles notifications for:
 * - Policy sharing (email/Teams notifications when policies are shared)
 * - Policy follows (notifications when followed policies are updated)
 * - Policy acknowledgments
 * - General policy notifications
 *
 * Uses Microsoft Graph API for email and Teams delivery
 * Works with PM_Notifications for in-app notifications
 */

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { MSGraphClientV3 } from '@microsoft/sp-http';
import { logger } from './LoggingService';
import { SystemLists, NotificationLists } from '../constants/SharePointListNames';

/**
 * Notification types for policy-related notifications
 */
export enum PolicyNotificationType {
  PolicyShared = 'PolicyShared',
  PolicyFollowed = 'PolicyFollowed',
  PolicyUpdated = 'PolicyUpdated',
  PolicyAcknowledgmentRequired = 'PolicyAcknowledgmentRequired',
  PolicyAcknowledged = 'PolicyAcknowledged',
  PolicyExpiring = 'PolicyExpiring',
  PolicyPublished = 'PolicyPublished',
  PolicyComment = 'PolicyComment',
  Custom = 'Custom'
}

/**
 * Notification channels
 */
export enum PolicyNotificationChannel {
  Email = 'Email',
  Teams = 'Teams',
  InApp = 'InApp',
  All = 'All'
}

/**
 * Notification priority levels
 */
export enum PolicyNotificationPriority {
  Low = 'Low',
  Normal = 'Normal',
  High = 'High',
  Urgent = 'Urgent'
}

/**
 * Notification status in SharePoint
 */
export enum NotificationQueueStatus {
  Pending = 'Pending',
  Processing = 'Processing',
  Sent = 'Sent',
  Failed = 'Failed',
  Retry = 'Retry'
}

/**
 * Notification queue item from PM_NotificationQueue
 */
export interface IPolicyNotificationQueueItem {
  Id: number;
  Title: string;
  NotificationType: PolicyNotificationType;
  RecipientEmail: string;
  RecipientUserId?: number;
  RecipientName?: string;
  SenderEmail?: string;
  SenderUserId?: number;
  SenderName?: string;
  PolicyId?: number;
  PolicyTitle?: string;
  PolicyVersion?: string;
  Message?: string;
  Channel: PolicyNotificationChannel;
  Priority: PolicyNotificationPriority;
  Status: NotificationQueueStatus;
  RetryCount: number;
  MaxRetries: number;
  LastError?: string;
  ScheduledSendTime?: Date;
  SentTime?: Date;
  RelatedShareId?: number;
  RelatedFollowId?: number;
  TeamsChannelId?: string;
  TeamsTeamId?: string;
  Created?: Date;
}

/**
 * Processor configuration
 */
export interface IPolicyNotificationProcessorConfig {
  enabled: boolean;
  intervalMs: number;
  maxBatchSize: number;
  defaultMaxRetries: number;
}

/**
 * Processing result
 */
export interface IProcessingResult {
  processed: number;
  sent: number;
  failed: number;
  retrying: number;
  errors: Array<{ notificationId: number; error: string }>;
}

/**
 * Service for processing policy notification queue
 */
export class PolicyNotificationQueueProcessor {
  private sp: SPFI;
  private context: WebPartContext;
  private siteUrl: string;
  private config: IPolicyNotificationProcessorConfig;
  private intervalHandle: NodeJS.Timeout | null = null;
  private isProcessing: boolean = false;

  private readonly QUEUE_LIST_NAME = SystemLists.NOTIFICATION_QUEUE;
  private readonly NOTIFICATIONS_LIST_NAME = NotificationLists.NOTIFICATIONS;

  constructor(sp: SPFI, context: WebPartContext, config?: Partial<IPolicyNotificationProcessorConfig>) {
    this.sp = sp;
    this.context = context;
    this.siteUrl = context.pageContext.web.absoluteUrl;
    this.config = {
      enabled: true,
      intervalMs: 60000, // 1 minute default
      maxBatchSize: 25,
      defaultMaxRetries: 3,
      ...config
    };
  }

  /**
   * Start the notification processor.
   * Verifies required lists exist before starting; silently disables if missing.
   */
  public start(): void {
    if (!this.config.enabled) {
      logger.info('PolicyNotificationQueueProcessor', 'Processor is disabled');
      return;
    }

    if (this.intervalHandle) {
      logger.warn('PolicyNotificationQueueProcessor', 'Processor already running');
      return;
    }

    // Verify required lists exist before starting the polling loop
    this.verifyListsExist().then(listsExist => {
      if (!listsExist) {
        logger.warn('PolicyNotificationQueueProcessor',
          `Required list '${this.QUEUE_LIST_NAME}' not found. ` +
          'Notification processing disabled until list is provisioned.');
        return;
      }

      logger.info('PolicyNotificationQueueProcessor', `Starting processor (interval: ${this.config.intervalMs}ms)`);

      // Run immediately once
      this.processQueue().catch(err => {
        logger.error('PolicyNotificationQueueProcessor', 'Error in initial processing run', err);
      });

      // Then schedule periodic runs
      this.intervalHandle = setInterval(() => {
        this.processQueue().catch(err => {
          logger.error('PolicyNotificationQueueProcessor', 'Error in scheduled processing run', err);
        });
      }, this.config.intervalMs);
    }).catch(err => {
      logger.warn('PolicyNotificationQueueProcessor', 'Failed to verify notification lists, processor disabled', err);
    });
  }

  /**
   * Check if the notification queue list exists
   */
  private async verifyListsExist(): Promise<boolean> {
    try {
      await this.sp.web.lists.getByTitle(this.QUEUE_LIST_NAME)();
      return true;
    } catch {
      return false;
    }
  }

  /**
   * Stop the notification processor
   */
  public stop(): void {
    if (this.intervalHandle) {
      clearInterval(this.intervalHandle);
      this.intervalHandle = null;
      logger.info('PolicyNotificationQueueProcessor', 'Processor stopped');
    }
  }

  /**
   * Process pending notifications in the queue
   */
  public async processQueue(): Promise<IProcessingResult> {
    if (this.isProcessing) {
      logger.info('PolicyNotificationQueueProcessor', 'Processing already in progress, skipping');
      return { processed: 0, sent: 0, failed: 0, retrying: 0, errors: [] };
    }

    this.isProcessing = true;
    const result: IProcessingResult = {
      processed: 0,
      sent: 0,
      failed: 0,
      retrying: 0,
      errors: []
    };

    try {
      // Get pending notifications (including those scheduled for now or earlier)
      const pendingItems = await this.getPendingNotifications();

      if (pendingItems.length === 0) {
        return result;
      }

      logger.info('PolicyNotificationQueueProcessor', `Processing ${pendingItems.length} pending notifications`);

      for (const item of pendingItems) {
        result.processed++;

        try {
          // Mark as processing
          await this.updateStatus(item.Id, NotificationQueueStatus.Processing);

          // Process based on channel
          const sendResult = await this.sendNotification(item);

          if (sendResult.success) {
            await this.markAsSent(item.Id);
            result.sent++;
          } else {
            const retryCount = (item.RetryCount || 0) + 1;
            const maxRetries = item.MaxRetries || this.config.defaultMaxRetries;

            if (retryCount >= maxRetries) {
              await this.markAsFailed(item.Id, sendResult.error || 'Max retries exceeded');
              result.failed++;
              result.errors.push({ notificationId: item.Id, error: sendResult.error || 'Max retries exceeded' });
            } else {
              await this.markForRetry(item.Id, retryCount, sendResult.error);
              result.retrying++;
            }
          }
        } catch (error) {
          const errorMessage = error instanceof Error ? error.message : String(error);
          result.failed++;
          result.errors.push({ notificationId: item.Id, error: errorMessage });
          logger.error('PolicyNotificationQueueProcessor', `Failed to process notification ${item.Id}`, error);

          // Mark as failed
          await this.markAsFailed(item.Id, errorMessage).catch(() => {});
        }
      }

      logger.info('PolicyNotificationQueueProcessor',
        `Processing complete: ${result.sent} sent, ${result.failed} failed, ${result.retrying} retrying`);

      return result;
    } finally {
      this.isProcessing = false;
    }
  }

  /**
   * Manually trigger processing
   */
  public async processSingle(notificationId: number): Promise<{ success: boolean; error?: string }> {
    try {
      const item = await this.sp.web.lists
        .getByTitle(this.QUEUE_LIST_NAME)
        .items.getById(notificationId)
        .select('*')();

      const notification = this.mapToNotification(item);
      return await this.sendNotification(notification);
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : String(error);
      return { success: false, error: errorMessage };
    }
  }

  /**
   * Get pending notifications from the queue
   */
  private async getPendingNotifications(): Promise<IPolicyNotificationQueueItem[]> {
    try {
      const now = new Date().toISOString();

      // Get pending and retry items that are due for processing
      const items = await this.sp.web.lists
        .getByTitle(this.QUEUE_LIST_NAME)
        .items
        .filter(`(Status eq 'Pending' or Status eq 'Retry') and (ScheduledSendTime eq null or ScheduledSendTime le datetime'${now}')`)
        .select('*')
        .orderBy('Priority', false) // High priority first
        .orderBy('Created', true)
        .top(this.config.maxBatchSize)();

      return items.map(item => this.mapToNotification(item));
    } catch (error) {
      logger.error('PolicyNotificationQueueProcessor', 'Failed to get pending notifications', error);
      return [];
    }
  }

  /**
   * Map SharePoint item to typed notification
   */
  private mapToNotification(item: Record<string, unknown>): IPolicyNotificationQueueItem {
    return {
      Id: item.Id as number,
      Title: item.Title as string,
      NotificationType: item.NotificationType as PolicyNotificationType,
      RecipientEmail: item.RecipientEmail as string,
      RecipientUserId: item.RecipientUserId as number | undefined,
      RecipientName: item.RecipientName as string | undefined,
      SenderEmail: item.SenderEmail as string | undefined,
      SenderUserId: item.SenderUserId as number | undefined,
      SenderName: item.SenderName as string | undefined,
      PolicyId: item.PolicyId as number | undefined,
      PolicyTitle: item.PolicyTitle as string | undefined,
      PolicyVersion: item.PolicyVersion as string | undefined,
      Message: item.Message as string | undefined,
      Channel: item.Channel as PolicyNotificationChannel,
      Priority: item.Priority as PolicyNotificationPriority,
      Status: item.Status as NotificationQueueStatus,
      RetryCount: (item.RetryCount as number) || 0,
      MaxRetries: (item.MaxRetries as number) || this.config.defaultMaxRetries,
      LastError: item.LastError as string | undefined,
      ScheduledSendTime: item.ScheduledSendTime ? new Date(item.ScheduledSendTime as string) : undefined,
      SentTime: item.SentTime ? new Date(item.SentTime as string) : undefined,
      RelatedShareId: item.RelatedShareId as number | undefined,
      RelatedFollowId: item.RelatedFollowId as number | undefined,
      TeamsChannelId: item.TeamsChannelId as string | undefined,
      TeamsTeamId: item.TeamsTeamId as string | undefined,
      Created: item.Created ? new Date(item.Created as string) : undefined
    };
  }

  /**
   * Send notification through the appropriate channel(s)
   */
  private async sendNotification(notification: IPolicyNotificationQueueItem): Promise<{ success: boolean; error?: string }> {
    const errors: string[] = [];
    let atLeastOneSent = false;

    try {
      const channel = notification.Channel || PolicyNotificationChannel.Email;

      // Send via Email
      if (channel === PolicyNotificationChannel.Email || channel === PolicyNotificationChannel.All) {
        const emailResult = await this.sendEmailNotification(notification);
        if (emailResult.success) {
          atLeastOneSent = true;
        } else if (emailResult.error) {
          errors.push(`Email: ${emailResult.error}`);
        }
      }

      // Send via Teams
      if (channel === PolicyNotificationChannel.Teams || channel === PolicyNotificationChannel.All) {
        if (notification.TeamsTeamId && notification.TeamsChannelId) {
          const teamsResult = await this.sendTeamsNotification(notification);
          if (teamsResult.success) {
            atLeastOneSent = true;
          } else if (teamsResult.error) {
            errors.push(`Teams: ${teamsResult.error}`);
          }
        }
      }

      // Send In-App notification
      if (channel === PolicyNotificationChannel.InApp || channel === PolicyNotificationChannel.All) {
        if (notification.RecipientUserId) {
          const inAppResult = await this.sendInAppNotification(notification);
          if (inAppResult.success) {
            atLeastOneSent = true;
          } else if (inAppResult.error) {
            errors.push(`InApp: ${inAppResult.error}`);
          }
        }
      }

      // Success if at least one channel succeeded
      if (atLeastOneSent) {
        return { success: true };
      }

      return { success: false, error: errors.join('; ') || 'No notifications sent' };
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : String(error);
      return { success: false, error: errorMessage };
    }
  }

  /**
   * Send email notification via Microsoft Graph API
   */
  private async sendEmailNotification(notification: IPolicyNotificationQueueItem): Promise<{ success: boolean; error?: string }> {
    try {
      if (!notification.RecipientEmail) {
        return { success: false, error: 'No recipient email address' };
      }

      const { subject, htmlBody } = this.buildEmailContent(notification);

      const emailMessage = {
        message: {
          subject,
          body: {
            contentType: 'HTML',
            content: htmlBody
          },
          toRecipients: [{
            emailAddress: { address: notification.RecipientEmail }
          }]
        },
        saveToSentItems: true
      };

      const graphClient: MSGraphClientV3 = await this.context.msGraphClientFactory.getClient('3');
      await graphClient.api('/me/sendMail').post(emailMessage);

      logger.info('PolicyNotificationQueueProcessor', `Email sent to ${notification.RecipientEmail}`);
      return { success: true };
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : String(error);
      logger.error('PolicyNotificationQueueProcessor', 'Failed to send email', error);
      return { success: false, error: errorMessage };
    }
  }

  /**
   * Send Teams notification via Microsoft Graph API
   */
  private async sendTeamsNotification(notification: IPolicyNotificationQueueItem): Promise<{ success: boolean; error?: string }> {
    try {
      if (!notification.TeamsTeamId || !notification.TeamsChannelId) {
        return { success: false, error: 'Teams team/channel not configured' };
      }

      const { subject, body } = this.buildTeamsContent(notification);

      const teamsMessage = {
        body: {
          contentType: 'html',
          content: `<strong>${subject}</strong><br/><br/>${body}`
        }
      };

      const graphClient: MSGraphClientV3 = await this.context.msGraphClientFactory.getClient('3');
      await graphClient
        .api(`/teams/${notification.TeamsTeamId}/channels/${notification.TeamsChannelId}/messages`)
        .post(teamsMessage);

      logger.info('PolicyNotificationQueueProcessor', `Teams message sent to channel ${notification.TeamsChannelId}`);
      return { success: true };
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : String(error);
      logger.error('PolicyNotificationQueueProcessor', 'Failed to send Teams notification', error);
      return { success: false, error: errorMessage };
    }
  }

  /**
   * Send in-app notification by creating item in PM_Notifications
   */
  private async sendInAppNotification(notification: IPolicyNotificationQueueItem): Promise<{ success: boolean; error?: string }> {
    try {
      if (!notification.RecipientUserId) {
        return { success: false, error: 'No recipient user ID for in-app notification' };
      }

      const { subject, body } = this.buildInAppContent(notification);

      const notificationData = {
        Title: subject,
        Message: body,
        RecipientId: notification.RecipientUserId,
        Type: this.mapNotificationTypeToInAppType(notification.NotificationType),
        Priority: notification.Priority || 'Normal',
        IsRead: false,
        RelatedItemType: 'Policy',
        RelatedItemId: notification.PolicyId,
        ActionUrl: this.buildActionUrl(notification)
      };

      await this.sp.web.lists
        .getByTitle(this.NOTIFICATIONS_LIST_NAME)
        .items.add(notificationData);

      logger.info('PolicyNotificationQueueProcessor', `In-app notification created for user ${notification.RecipientUserId}`);
      return { success: true };
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : String(error);
      logger.error('PolicyNotificationQueueProcessor', 'Failed to create in-app notification', error);
      return { success: false, error: errorMessage };
    }
  }

  /**
   * Build email content based on notification type
   */
  private buildEmailContent(notification: IPolicyNotificationQueueItem): { subject: string; htmlBody: string } {
    const { subject, body } = this.buildNotificationText(notification);
    const policyUrl = notification.PolicyId
      ? `${this.siteUrl}/SitePages/PolicyDetails.aspx?policyId=${notification.PolicyId}`
      : this.siteUrl;

    const colors = this.getNotificationColors(notification.NotificationType);

    const htmlBody = `
<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <style>
    body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 0; padding: 0; background-color: #f3f2f1; }
    .container { max-width: 600px; margin: 0 auto; padding: 20px; }
    .card { background: #fff; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); overflow: hidden; }
    .header { background: ${colors.bg}; border-left: 4px solid ${colors.border}; padding: 16px 20px; }
    .header h1 { margin: 0; font-size: 18px; color: ${colors.text}; }
    .content { padding: 20px; }
    .content p { margin: 0 0 16px; color: #323130; line-height: 1.5; }
    .policy-card { background: #faf9f8; border: 1px solid #edebe9; border-radius: 4px; padding: 16px; margin: 16px 0; }
    .policy-title { font-size: 16px; font-weight: 600; color: #323130; margin-bottom: 8px; }
    .policy-meta { font-size: 13px; color: #605e5c; }
    .sender-info { margin-top: 16px; padding-top: 16px; border-top: 1px solid #edebe9; font-size: 13px; color: #605e5c; }
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
        ${notification.PolicyTitle ? `
        <div class="policy-card">
          <div class="policy-title">${notification.PolicyTitle}</div>
          <div class="policy-meta">
            ${notification.PolicyVersion ? `Version: ${notification.PolicyVersion}` : ''}
          </div>
        </div>
        ` : ''}
        ${notification.Message ? `<p><em>"${notification.Message}"</em></p>` : ''}
        ${notification.SenderName ? `
        <div class="sender-info">
          Shared by: <strong>${notification.SenderName}</strong>${notification.SenderEmail ? ` (${notification.SenderEmail})` : ''}
        </div>
        ` : ''}
        <a href="${policyUrl}" class="button">View Policy</a>
      </div>
      <div class="footer">
        This is an automated notification from the DWx Policy Management System.
      </div>
    </div>
  </div>
</body>
</html>`;

    return { subject, htmlBody };
  }

  /**
   * Build Teams content
   */
  private buildTeamsContent(notification: IPolicyNotificationQueueItem): { subject: string; body: string } {
    return this.buildNotificationText(notification);
  }

  /**
   * Build in-app notification content
   */
  private buildInAppContent(notification: IPolicyNotificationQueueItem): { subject: string; body: string } {
    return this.buildNotificationText(notification);
  }

  /**
   * Build notification text based on type
   */
  private buildNotificationText(notification: IPolicyNotificationQueueItem): { subject: string; body: string } {
    const senderName = notification.SenderName || 'Someone';
    const policyTitle = notification.PolicyTitle || 'a policy';

    let subject = '';
    let body = '';

    switch (notification.NotificationType) {
      case PolicyNotificationType.PolicyShared:
        subject = `Policy Shared: ${policyTitle}`;
        body = `${senderName} has shared the policy "${policyTitle}" with you.`;
        break;

      case PolicyNotificationType.PolicyFollowed:
        subject = `Now Following: ${policyTitle}`;
        body = `You are now following the policy "${policyTitle}". You will receive notifications when this policy is updated.`;
        break;

      case PolicyNotificationType.PolicyUpdated:
        subject = `Policy Updated: ${policyTitle}`;
        body = `The policy "${policyTitle}" has been updated${notification.PolicyVersion ? ` to version ${notification.PolicyVersion}` : ''}.`;
        break;

      case PolicyNotificationType.PolicyAcknowledgmentRequired:
        subject = `Action Required: Acknowledge ${policyTitle}`;
        body = `You are required to acknowledge the policy "${policyTitle}". Please review and acknowledge at your earliest convenience.`;
        break;

      case PolicyNotificationType.PolicyAcknowledged:
        subject = `Policy Acknowledged: ${policyTitle}`;
        body = `Thank you for acknowledging the policy "${policyTitle}".`;
        break;

      case PolicyNotificationType.PolicyExpiring:
        subject = `Policy Expiring: ${policyTitle}`;
        body = `The policy "${policyTitle}" is approaching its expiration date and may require review.`;
        break;

      case PolicyNotificationType.PolicyPublished:
        subject = `New Policy Published: ${policyTitle}`;
        body = `A new policy "${policyTitle}" has been published and is now available for review.`;
        break;

      case PolicyNotificationType.PolicyComment:
        subject = `New Comment on: ${policyTitle}`;
        body = `${senderName} has commented on the policy "${policyTitle}".`;
        break;

      case PolicyNotificationType.Custom:
      default:
        subject = notification.Title || 'Policy Notification';
        body = notification.Message || 'You have a new policy notification.';
    }

    // Append custom message if present and not already used
    if (notification.Message && notification.NotificationType !== PolicyNotificationType.Custom) {
      body += ` Note: "${notification.Message}"`;
    }

    return { subject, body };
  }

  /**
   * Get notification colors based on type
   */
  private getNotificationColors(type: PolicyNotificationType): { bg: string; border: string; text: string } {
    switch (type) {
      case PolicyNotificationType.PolicyAcknowledgmentRequired:
      case PolicyNotificationType.PolicyExpiring:
        return { bg: '#fff4ce', border: '#ff8c00', text: '#8a6914' };

      case PolicyNotificationType.PolicyAcknowledged:
      case PolicyNotificationType.PolicyPublished:
        return { bg: '#dff6dd', border: '#107c10', text: '#0b5c0b' };

      case PolicyNotificationType.PolicyShared:
      case PolicyNotificationType.PolicyFollowed:
      case PolicyNotificationType.PolicyUpdated:
      case PolicyNotificationType.PolicyComment:
      default:
        return { bg: '#e7f3ff', border: '#0078d4', text: '#004578' };
    }
  }

  /**
   * Map policy notification type to in-app notification type
   */
  private mapNotificationTypeToInAppType(type: PolicyNotificationType): string {
    switch (type) {
      case PolicyNotificationType.PolicyShared:
        return 'PolicyShare';
      case PolicyNotificationType.PolicyFollowed:
        return 'PolicyFollow';
      case PolicyNotificationType.PolicyUpdated:
        return 'PolicyUpdate';
      case PolicyNotificationType.PolicyAcknowledgmentRequired:
        return 'PolicyAcknowledgment';
      case PolicyNotificationType.PolicyExpiring:
        return 'PolicyExpiring';
      default:
        return 'Policy';
    }
  }

  /**
   * Build action URL for the notification
   */
  private buildActionUrl(notification: IPolicyNotificationQueueItem): string {
    if (notification.PolicyId) {
      return `${this.siteUrl}/SitePages/PolicyDetails.aspx?policyId=${notification.PolicyId}`;
    }
    return `${this.siteUrl}/SitePages/PolicyHub.aspx`;
  }

  /**
   * Update notification status
   */
  private async updateStatus(notificationId: number, status: NotificationQueueStatus): Promise<void> {
    await this.sp.web.lists
      .getByTitle(this.QUEUE_LIST_NAME)
      .items.getById(notificationId)
      .update({ Status: status });
  }

  /**
   * Mark notification as sent
   */
  private async markAsSent(notificationId: number): Promise<void> {
    await this.sp.web.lists
      .getByTitle(this.QUEUE_LIST_NAME)
      .items.getById(notificationId)
      .update({
        Status: NotificationQueueStatus.Sent,
        SentTime: new Date().toISOString()
      });
  }

  /**
   * Mark notification as failed
   */
  private async markAsFailed(notificationId: number, error: string): Promise<void> {
    await this.sp.web.lists
      .getByTitle(this.QUEUE_LIST_NAME)
      .items.getById(notificationId)
      .update({
        Status: NotificationQueueStatus.Failed,
        LastError: error.substring(0, 255)
      });
  }

  /**
   * Mark notification for retry
   */
  private async markForRetry(notificationId: number, retryCount: number, error?: string): Promise<void> {
    const updateData: Record<string, unknown> = {
      Status: NotificationQueueStatus.Retry,
      RetryCount: retryCount
    };

    if (error) {
      updateData.LastError = error.substring(0, 255);
    }

    await this.sp.web.lists
      .getByTitle(this.QUEUE_LIST_NAME)
      .items.getById(notificationId)
      .update(updateData);
  }

  /**
   * Get queue statistics
   */
  public async getStats(): Promise<{
    pending: number;
    processing: number;
    sent: number;
    failed: number;
    retry: number;
    isRunning: boolean;
    isProcessing: boolean;
  }> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(this.QUEUE_LIST_NAME)
        .items
        .select('Status')();

      const stats = {
        pending: 0,
        processing: 0,
        sent: 0,
        failed: 0,
        retry: 0,
        isRunning: this.intervalHandle !== null,
        isProcessing: this.isProcessing
      };

      items.forEach(item => {
        switch (item.Status) {
          case NotificationQueueStatus.Pending:
            stats.pending++;
            break;
          case NotificationQueueStatus.Processing:
            stats.processing++;
            break;
          case NotificationQueueStatus.Sent:
            stats.sent++;
            break;
          case NotificationQueueStatus.Failed:
            stats.failed++;
            break;
          case NotificationQueueStatus.Retry:
            stats.retry++;
            break;
        }
      });

      return stats;
    } catch (error) {
      logger.error('PolicyNotificationQueueProcessor', 'Failed to get stats', error);
      return {
        pending: 0,
        processing: 0,
        sent: 0,
        failed: 0,
        retry: 0,
        isRunning: this.intervalHandle !== null,
        isProcessing: this.isProcessing
      };
    }
  }

  /**
   * Retry all failed notifications
   */
  public async retryFailed(): Promise<number> {
    try {
      const failedItems = await this.sp.web.lists
        .getByTitle(this.QUEUE_LIST_NAME)
        .items
        .filter(`Status eq 'Failed'`)
        .select('Id')
        .top(50)();

      for (const item of failedItems) {
        await this.sp.web.lists
          .getByTitle(this.QUEUE_LIST_NAME)
          .items.getById(item.Id)
          .update({
            Status: NotificationQueueStatus.Pending,
            RetryCount: 0,
            LastError: null
          });
      }

      logger.info('PolicyNotificationQueueProcessor', `Reset ${failedItems.length} failed notifications for retry`);
      return failedItems.length;
    } catch (error) {
      logger.error('PolicyNotificationQueueProcessor', 'Failed to retry failed notifications', error);
      return 0;
    }
  }

  /**
   * Clear old sent notifications (cleanup)
   */
  public async cleanupOldNotifications(daysOld: number = 30): Promise<number> {
    try {
      const cutoffDate = new Date();
      cutoffDate.setDate(cutoffDate.getDate() - daysOld);

      const oldItems = await this.sp.web.lists
        .getByTitle(this.QUEUE_LIST_NAME)
        .items
        .filter(`Status eq 'Sent' and SentTime lt datetime'${cutoffDate.toISOString()}'`)
        .select('Id')
        .top(100)();

      for (const item of oldItems) {
        await this.sp.web.lists
          .getByTitle(this.QUEUE_LIST_NAME)
          .items.getById(item.Id)
          .delete();
      }

      logger.info('PolicyNotificationQueueProcessor', `Cleaned up ${oldItems.length} old notifications`);
      return oldItems.length;
    } catch (error) {
      logger.error('PolicyNotificationQueueProcessor', 'Failed to cleanup old notifications', error);
      return 0;
    }
  }
}
