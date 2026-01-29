// @ts-nocheck
/* eslint-disable */
/**
 * NotificationProcessorService
 * Scheduled processor for handling notification delivery
 * CRITICAL: Ensures notifications are delivered even when UI polling is not active
 */

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import { logger } from '../LoggingService';
import {
  retryWithDLQ,
  notificationDLQ,
  NOTIFICATION_RETRY_OPTIONS
} from '../../utils/retryUtils';
import { Priority } from '../../models';

/**
 * Notification status in SharePoint
 */
export enum NotificationDeliveryStatus {
  Pending = 'Pending',
  Processing = 'Processing',
  Delivered = 'Delivered',
  Failed = 'Failed',
  Expired = 'Expired'
}

/**
 * Notification item from SharePoint
 */
export interface INotificationItem {
  Id: number;
  Title: string;
  Message: string;
  RecipientId: number;
  NotificationType: string;
  Priority: Priority;
  LinkUrl?: string;
  ProcessId?: string;
  IsRead: boolean;
  SentDate?: Date;
  DeliveryStatus?: NotificationDeliveryStatus;
  DeliveryAttempts?: number;
  ExpirationDate?: Date;
}

/**
 * Processor configuration
 */
export interface INotificationProcessorConfig {
  enabled: boolean;
  intervalMs: number;
  maxBatchSize: number;
  maxRetries: number;
  expireAfterDays: number;
}

/**
 * Processing result
 */
export interface INotificationProcessingResult {
  processed: number;
  delivered: number;
  failed: number;
  expired: number;
  errors: Array<{ notificationId: number; error: string }>;
}

export class NotificationProcessorService {
  private sp: SPFI;
  private config: INotificationProcessorConfig;
  private intervalHandle: any = null;
  private isProcessing: boolean = false;
  private readonly LIST_NAME = 'PM_Notifications';

  constructor(sp: SPFI, config?: Partial<INotificationProcessorConfig>) {
    this.sp = sp;
    this.config = {
      enabled: true,
      intervalMs: 60000, // 1 minute default
      maxBatchSize: 50,
      maxRetries: 3,
      expireAfterDays: 7,
      ...config
    };
  }

  /**
   * Start the notification processor
   */
  public start(): void {
    if (!this.config.enabled) {
      logger.info('NotificationProcessor', 'Notification processor is disabled');
      return;
    }

    if (this.intervalHandle) {
      logger.warn('NotificationProcessor', 'Processor already running');
      return;
    }

    logger.info('NotificationProcessor', `Starting notification processor (interval: ${this.config.intervalMs}ms)`);

    // Run immediately once
    this.processNotifications().catch(err => {
      logger.error('NotificationProcessor', 'Error in initial processing run', err);
    });

    // Then schedule periodic runs
    this.intervalHandle = setInterval(() => {
      this.processNotifications().catch(err => {
        logger.error('NotificationProcessor', 'Error in scheduled processing run', err);
      });
    }, this.config.intervalMs);
  }

  /**
   * Stop the notification processor
   */
  public stop(): void {
    if (this.intervalHandle) {
      clearInterval(this.intervalHandle);
      this.intervalHandle = null;
      logger.info('NotificationProcessor', 'Notification processor stopped');
    }
  }

  /**
   * Process pending notifications
   */
  public async processNotifications(): Promise<INotificationProcessingResult> {
    if (this.isProcessing) {
      logger.info('NotificationProcessor', 'Processing already in progress, skipping');
      return { processed: 0, delivered: 0, failed: 0, expired: 0, errors: [] };
    }

    this.isProcessing = true;
    const result: INotificationProcessingResult = {
      processed: 0,
      delivered: 0,
      failed: 0,
      expired: 0,
      errors: []
    };

    try {
      // Get pending notifications
      const pendingNotifications = await this.getPendingNotifications();

      if (pendingNotifications.length === 0) {
        return result;
      }

      logger.info('NotificationProcessor', `Processing ${pendingNotifications.length} pending notifications`);

      for (const notification of pendingNotifications) {
        result.processed++;

        try {
          // Check if expired
          if (notification.ExpirationDate && new Date(notification.ExpirationDate) < new Date()) {
            await this.markExpired(notification.Id);
            result.expired++;
            continue;
          }

          // Mark as processing
          await this.markProcessing(notification.Id);

          // Deliver the notification
          const deliveryResult = await this.deliverNotification(notification);

          if (deliveryResult.success) {
            await this.markDelivered(notification.Id);
            result.delivered++;
          } else {
            const attempts = (notification.DeliveryAttempts || 0) + 1;
            if (attempts >= this.config.maxRetries) {
              await this.markFailed(notification.Id, deliveryResult.error || 'Max retries exceeded');
              result.failed++;
              result.errors.push({ notificationId: notification.Id, error: deliveryResult.error || 'Max retries exceeded' });
            } else {
              await this.updateAttempts(notification.Id, attempts, deliveryResult.error);
            }
          }
        } catch (error) {
          const errorMessage = error instanceof Error ? error.message : String(error);
          result.failed++;
          result.errors.push({ notificationId: notification.Id, error: errorMessage });
          logger.error('NotificationProcessor', `Failed to process notification ${notification.Id}`, error);
        }
      }

      // Process expired notifications
      const expiredCount = await this.processExpiredNotifications();
      result.expired += expiredCount;

      logger.info('NotificationProcessor',
        `Processing complete: ${result.delivered} delivered, ${result.failed} failed, ${result.expired} expired`);

      return result;
    } finally {
      this.isProcessing = false;
    }
  }

  /**
   * Get pending notifications that need to be processed
   */
  private async getPendingNotifications(): Promise<INotificationItem[]> {
    try {
      const items = await this.sp.web.lists.getByTitle(this.LIST_NAME)
        .items
        .filter(`DeliveryStatus eq 'Pending' or DeliveryStatus eq null`)
        .select('Id', 'Title', 'Message', 'RecipientId', 'NotificationType', 'Priority',
                'LinkUrl', 'ProcessId', 'IsRead', 'SentDate', 'DeliveryStatus',
                'DeliveryAttempts', 'ExpirationDate')
        .orderBy('SentDate', true)
        .top(this.config.maxBatchSize)();

      return items.map(item => ({
        ...item,
        SentDate: item.SentDate ? new Date(item.SentDate) : undefined,
        ExpirationDate: item.ExpirationDate ? new Date(item.ExpirationDate) : undefined
      }));
    } catch (error) {
      logger.error('NotificationProcessor', 'Failed to get pending notifications', error);
      return [];
    }
  }

  /**
   * Deliver a notification (in this implementation, just mark as sent)
   * In a full implementation, this could send email, Teams messages, etc.
   */
  private async deliverNotification(notification: INotificationItem): Promise<{ success: boolean; error?: string }> {
    return await retryWithDLQ(
      async () => {
        // The notification is already in SharePoint, so "delivery" means:
        // 1. The item exists and is accessible
        // 2. Could extend this to send Teams/email notifications

        // For now, simply verify the item is properly formed
        if (!notification.RecipientId) {
          throw new Error('Notification has no recipient');
        }

        if (!notification.Message) {
          throw new Error('Notification has no message');
        }

        // Could add Teams/email delivery here:
        // await this.sendTeamsNotification(notification);
        // await this.sendEmailNotification(notification);

        return true;
      },
      'notification-delivery',
      { notificationId: notification.Id, recipientId: notification.RecipientId },
      { ...NOTIFICATION_RETRY_OPTIONS, maxRetries: 1 }, // Single retry per processing cycle
      notificationDLQ,
      { source: 'NotificationProcessor' }
    ).then(result => ({
      success: result.success,
      error: result.error?.message
    }));
  }

  /**
   * Mark notification as processing
   */
  private async markProcessing(notificationId: number): Promise<void> {
    await this.sp.web.lists.getByTitle(this.LIST_NAME)
      .items.getById(notificationId)
      .update({
        DeliveryStatus: NotificationDeliveryStatus.Processing
      });
  }

  /**
   * Mark notification as delivered
   */
  private async markDelivered(notificationId: number): Promise<void> {
    await this.sp.web.lists.getByTitle(this.LIST_NAME)
      .items.getById(notificationId)
      .update({
        DeliveryStatus: NotificationDeliveryStatus.Delivered,
        SentDate: new Date().toISOString()
      });
  }

  /**
   * Mark notification as failed
   */
  private async markFailed(notificationId: number, error: string): Promise<void> {
    await this.sp.web.lists.getByTitle(this.LIST_NAME)
      .items.getById(notificationId)
      .update({
        DeliveryStatus: NotificationDeliveryStatus.Failed,
        DeliveryError: error.substring(0, 255) // Truncate to fit in SharePoint
      });
  }

  /**
   * Mark notification as expired
   */
  private async markExpired(notificationId: number): Promise<void> {
    await this.sp.web.lists.getByTitle(this.LIST_NAME)
      .items.getById(notificationId)
      .update({
        DeliveryStatus: NotificationDeliveryStatus.Expired
      });
  }

  /**
   * Update delivery attempts
   */
  private async updateAttempts(notificationId: number, attempts: number, error?: string): Promise<void> {
    const updateData: any = {
      DeliveryStatus: NotificationDeliveryStatus.Pending, // Reset to pending for next cycle
      DeliveryAttempts: attempts
    };

    if (error) {
      updateData.DeliveryError = error.substring(0, 255);
    }

    await this.sp.web.lists.getByTitle(this.LIST_NAME)
      .items.getById(notificationId)
      .update(updateData);
  }

  /**
   * Process notifications that have expired
   */
  private async processExpiredNotifications(): Promise<number> {
    try {
      const expirationDate = new Date();
      expirationDate.setDate(expirationDate.getDate() - this.config.expireAfterDays);

      const expiredItems = await this.sp.web.lists.getByTitle(this.LIST_NAME)
        .items
        .filter(`(DeliveryStatus eq 'Pending' or DeliveryStatus eq null) and SentDate lt datetime'${expirationDate.toISOString()}'`)
        .select('Id')
        .top(100)();

      for (const item of expiredItems) {
        await this.markExpired(item.Id);
      }

      return expiredItems.length;
    } catch (error) {
      logger.warn('NotificationProcessor', 'Failed to process expired notifications', error);
      return 0;
    }
  }

  /**
   * Get processor statistics
   */
  public async getStats(): Promise<{
    pending: number;
    delivered: number;
    failed: number;
    expired: number;
    isRunning: boolean;
    isProcessing: boolean;
  }> {
    try {
      const items = await this.sp.web.lists.getByTitle(this.LIST_NAME)
        .items
        .select('DeliveryStatus')();

      const stats = {
        pending: 0,
        delivered: 0,
        failed: 0,
        expired: 0,
        isRunning: this.intervalHandle !== null,
        isProcessing: this.isProcessing
      };

      items.forEach(item => {
        switch (item.DeliveryStatus) {
          case NotificationDeliveryStatus.Pending:
          case null:
          case undefined:
            stats.pending++;
            break;
          case NotificationDeliveryStatus.Delivered:
            stats.delivered++;
            break;
          case NotificationDeliveryStatus.Failed:
            stats.failed++;
            break;
          case NotificationDeliveryStatus.Expired:
            stats.expired++;
            break;
        }
      });

      return stats;
    } catch (error) {
      logger.error('NotificationProcessor', 'Failed to get stats', error);
      return {
        pending: 0,
        delivered: 0,
        failed: 0,
        expired: 0,
        isRunning: this.intervalHandle !== null,
        isProcessing: this.isProcessing
      };
    }
  }

  /**
   * Manually trigger notification retry for failed notifications
   */
  public async retryFailedNotifications(): Promise<number> {
    try {
      const failedItems = await this.sp.web.lists.getByTitle(this.LIST_NAME)
        .items
        .filter(`DeliveryStatus eq 'Failed'`)
        .select('Id')
        .top(50)();

      for (const item of failedItems) {
        await this.sp.web.lists.getByTitle(this.LIST_NAME)
          .items.getById(item.Id)
          .update({
            DeliveryStatus: NotificationDeliveryStatus.Pending,
            DeliveryAttempts: 0
          });
      }

      logger.info('NotificationProcessor', `Reset ${failedItems.length} failed notifications for retry`);
      return failedItems.length;
    } catch (error) {
      logger.error('NotificationProcessor', 'Failed to retry failed notifications', error);
      return 0;
    }
  }
}
