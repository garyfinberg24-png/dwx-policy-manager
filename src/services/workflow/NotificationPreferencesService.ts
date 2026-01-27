// @ts-nocheck
/* eslint-disable */
/**
 * NotificationPreferencesService
 * Manages user notification preferences for the JML workflow system
 *
 * Features:
 * - Per-user notification channel preferences (Email, Teams, InApp)
 * - Event type subscriptions
 * - Digest/summary options
 * - Quiet hours configuration
 */

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';

import { NotificationType, Priority } from '../../models/ICommon';
import { logger } from '../LoggingService';

// ============================================================================
// INTERFACES
// ============================================================================

/**
 * Notification channel options
 */
export enum NotificationChannel {
  Email = 'Email',
  Teams = 'Teams',
  InApp = 'InApp',
  SMS = 'SMS'
}

/**
 * Digest frequency options
 */
export enum DigestFrequency {
  Immediate = 'Immediate',
  Hourly = 'Hourly',
  Daily = 'Daily',
  Weekly = 'Weekly',
  None = 'None'
}

/**
 * Notification event types that users can configure
 */
export enum NotificationEventType {
  // Task Events
  TaskAssigned = 'TaskAssigned',
  TaskDueSoon = 'TaskDueSoon',
  TaskOverdue = 'TaskOverdue',
  TaskCompleted = 'TaskCompleted',
  TaskUnblocked = 'TaskUnblocked',

  // Approval Events
  ApprovalRequired = 'ApprovalRequired',
  ApprovalCompleted = 'ApprovalCompleted',
  ApprovalRejected = 'ApprovalRejected',
  ApprovalEscalated = 'ApprovalEscalated',

  // Process Events
  ProcessStarted = 'ProcessStarted',
  ProcessCompleted = 'ProcessCompleted',
  ProcessBlocked = 'ProcessBlocked',

  // Workflow Events
  WorkflowStepComplete = 'WorkflowStepComplete',
  WorkflowError = 'WorkflowError',

  // System Events
  SystemAnnouncement = 'SystemAnnouncement',
  Reminder = 'Reminder'
}

/**
 * Channel preference for a specific event type
 */
export interface IEventChannelPreference {
  eventType: NotificationEventType;
  channels: NotificationChannel[];
  enabled: boolean;
  digestFrequency: DigestFrequency;
}

/**
 * Quiet hours configuration
 */
export interface IQuietHours {
  enabled: boolean;
  startTime: string;  // HH:mm format
  endTime: string;    // HH:mm format
  timezone: string;
  excludeUrgent: boolean;  // Allow urgent notifications during quiet hours
  excludeDays: number[];   // 0 = Sunday, 6 = Saturday
}

/**
 * User notification preferences
 */
export interface IUserNotificationPreferences {
  id?: number;
  userId: number;
  userEmail: string;

  // Global settings
  globalEnabled: boolean;
  defaultChannels: NotificationChannel[];
  defaultDigestFrequency: DigestFrequency;

  // Per-event preferences
  eventPreferences: IEventChannelPreference[];

  // Quiet hours
  quietHours: IQuietHours;

  // Metadata
  createdDate: Date;
  modifiedDate: Date;
}

/**
 * Resolved notification delivery settings for a specific notification
 */
export interface IResolvedDeliverySettings {
  userId: number;
  userEmail: string;
  shouldDeliver: boolean;
  channels: NotificationChannel[];
  isDigest: boolean;
  digestFrequency: DigestFrequency;
  reason?: string;  // Why notification was blocked/modified
}

/**
 * Digest item for batching notifications
 */
export interface IDigestItem {
  id?: number;
  userId: number;
  eventType: NotificationEventType;
  subject: string;
  body: string;
  priority: Priority;
  relatedItemId?: number;
  relatedItemType?: string;
  scheduledDigestTime: Date;
  createdDate: Date;
}

// ============================================================================
// SERVICE IMPLEMENTATION
// ============================================================================

export class NotificationPreferencesService {
  private sp: SPFI;
  private readonly PREFERENCES_LIST = 'JML_UserPreferences';
  private readonly DIGEST_QUEUE_LIST = 'JML_DigestQueue';

  // Cache for preferences to reduce list queries
  private preferencesCache: Map<number, { preferences: IUserNotificationPreferences; expiry: Date }> = new Map();
  private readonly CACHE_TTL_MINUTES = 15;

  constructor(sp: SPFI) {
    this.sp = sp;
  }

  // ============================================================================
  // PREFERENCE MANAGEMENT
  // ============================================================================

  /**
   * Get notification preferences for a user
   */
  public async getUserPreferences(userId: number): Promise<IUserNotificationPreferences | null> {
    try {
      // Check cache first
      const cached = this.preferencesCache.get(userId);
      if (cached && cached.expiry > new Date()) {
        return cached.preferences;
      }

      const items = await this.sp.web.lists.getByTitle(this.PREFERENCES_LIST)
        .items
        .filter(`UserId eq ${userId}`)
        .top(1)();

      if (items.length === 0) {
        return null;
      }

      const item = items[0];
      const preferences = this.mapItemToPreferences(item);

      // Update cache
      this.updateCache(userId, preferences);

      return preferences;
    } catch (error) {
      logger.error('NotificationPreferencesService', `Error getting preferences for user ${userId}`, error);
      return null;
    }
  }

  /**
   * Get or create default preferences for a user
   */
  public async getOrCreatePreferences(userId: number, userEmail: string): Promise<IUserNotificationPreferences> {
    const existing = await this.getUserPreferences(userId);

    if (existing) {
      return existing;
    }

    // Create default preferences
    return this.createDefaultPreferences(userId, userEmail);
  }

  /**
   * Create default notification preferences for a new user
   */
  public async createDefaultPreferences(userId: number, userEmail: string): Promise<IUserNotificationPreferences> {
    try {
      const defaultPreferences: IUserNotificationPreferences = {
        userId,
        userEmail,
        globalEnabled: true,
        defaultChannels: [NotificationChannel.Email, NotificationChannel.InApp],
        defaultDigestFrequency: DigestFrequency.Immediate,
        eventPreferences: this.getDefaultEventPreferences(),
        quietHours: {
          enabled: false,
          startTime: '22:00',
          endTime: '07:00',
          timezone: 'UTC',
          excludeUrgent: true,
          excludeDays: [0, 6]  // Weekend
        },
        createdDate: new Date(),
        modifiedDate: new Date()
      };

      const result = await this.sp.web.lists.getByTitle(this.PREFERENCES_LIST)
        .items.add({
          UserId: userId,
          UserEmail: userEmail,
          GlobalEnabled: true,
          DefaultChannels: JSON.stringify(defaultPreferences.defaultChannels),
          DefaultDigestFrequency: defaultPreferences.defaultDigestFrequency,
          EventPreferences: JSON.stringify(defaultPreferences.eventPreferences),
          QuietHours: JSON.stringify(defaultPreferences.quietHours),
          CreatedDate: new Date().toISOString(),
          ModifiedDate: new Date().toISOString()
        });

      defaultPreferences.id = result.data.Id;

      // Update cache
      this.updateCache(userId, defaultPreferences);

      logger.info('NotificationPreferencesService',
        `Created default preferences for user ${userId}`);

      return defaultPreferences;
    } catch (error) {
      logger.error('NotificationPreferencesService', 'Error creating default preferences', error);
      throw error;
    }
  }

  /**
   * Update user notification preferences
   */
  public async updatePreferences(
    userId: number,
    updates: Partial<IUserNotificationPreferences>
  ): Promise<IUserNotificationPreferences | null> {
    try {
      const existing = await this.getUserPreferences(userId);
      if (!existing || !existing.id) {
        logger.warn('NotificationPreferencesService',
          `No preferences found for user ${userId}`);
        return null;
      }

      const updatePayload: Record<string, unknown> = {
        ModifiedDate: new Date().toISOString()
      };

      if (updates.globalEnabled !== undefined) {
        updatePayload.GlobalEnabled = updates.globalEnabled;
      }

      if (updates.defaultChannels) {
        updatePayload.DefaultChannels = JSON.stringify(updates.defaultChannels);
      }

      if (updates.defaultDigestFrequency) {
        updatePayload.DefaultDigestFrequency = updates.defaultDigestFrequency;
      }

      if (updates.eventPreferences) {
        updatePayload.EventPreferences = JSON.stringify(updates.eventPreferences);
      }

      if (updates.quietHours) {
        updatePayload.QuietHours = JSON.stringify(updates.quietHours);
      }

      await this.sp.web.lists.getByTitle(this.PREFERENCES_LIST)
        .items.getById(existing.id)
        .update(updatePayload);

      // Invalidate cache
      this.preferencesCache.delete(userId);

      // Get and return updated preferences
      return this.getUserPreferences(userId);
    } catch (error) {
      logger.error('NotificationPreferencesService', 'Error updating preferences', error);
      return null;
    }
  }

  /**
   * Update a specific event preference
   */
  public async updateEventPreference(
    userId: number,
    eventType: NotificationEventType,
    preference: Partial<IEventChannelPreference>
  ): Promise<boolean> {
    try {
      const prefs = await this.getUserPreferences(userId);
      if (!prefs) return false;

      const eventPrefs = [...prefs.eventPreferences];
      const existingIndex = eventPrefs.findIndex(e => e.eventType === eventType);

      if (existingIndex >= 0) {
        eventPrefs[existingIndex] = {
          ...eventPrefs[existingIndex],
          ...preference
        };
      } else {
        eventPrefs.push({
          eventType,
          channels: preference.channels || prefs.defaultChannels,
          enabled: preference.enabled !== undefined ? preference.enabled : true,
          digestFrequency: preference.digestFrequency || prefs.defaultDigestFrequency
        });
      }

      await this.updatePreferences(userId, { eventPreferences: eventPrefs });
      return true;
    } catch (error) {
      logger.error('NotificationPreferencesService', 'Error updating event preference', error);
      return false;
    }
  }

  // ============================================================================
  // DELIVERY RESOLUTION
  // ============================================================================

  /**
   * Resolve delivery settings for a notification
   * Determines if and how a notification should be delivered based on user preferences
   */
  public async resolveDeliverySettings(
    userId: number,
    userEmail: string,
    eventType: NotificationEventType,
    priority: Priority = Priority.Medium
  ): Promise<IResolvedDeliverySettings> {
    try {
      const prefs = await this.getOrCreatePreferences(userId, userEmail);

      // Check if globally disabled
      if (!prefs.globalEnabled) {
        return {
          userId,
          userEmail,
          shouldDeliver: false,
          channels: [],
          isDigest: false,
          digestFrequency: DigestFrequency.None,
          reason: 'User has disabled all notifications'
        };
      }

      // Find event-specific preferences
      const eventPref = prefs.eventPreferences.find(e => e.eventType === eventType);

      // Check if event type is disabled
      if (eventPref && !eventPref.enabled) {
        return {
          userId,
          userEmail,
          shouldDeliver: false,
          channels: [],
          isDigest: false,
          digestFrequency: DigestFrequency.None,
          reason: `User has disabled ${eventType} notifications`
        };
      }

      // Determine channels
      const channels = eventPref?.channels || prefs.defaultChannels;

      // Determine digest frequency
      const digestFrequency = eventPref?.digestFrequency || prefs.defaultDigestFrequency;
      const isDigest = digestFrequency !== DigestFrequency.Immediate;

      // Check quiet hours
      if (prefs.quietHours.enabled && this.isInQuietHours(prefs.quietHours)) {
        // During quiet hours, only deliver critical notifications if configured
        if (priority !== Priority.Critical || !prefs.quietHours.excludeUrgent) {
          return {
            userId,
            userEmail,
            shouldDeliver: true,
            channels,
            isDigest: true,  // Queue for later delivery
            digestFrequency: DigestFrequency.Daily,  // Deliver after quiet hours
            reason: 'Queued due to quiet hours'
          };
        }
      }

      return {
        userId,
        userEmail,
        shouldDeliver: true,
        channels,
        isDigest,
        digestFrequency
      };
    } catch (error) {
      logger.error('NotificationPreferencesService', 'Error resolving delivery settings', error);

      // Default to delivering via all channels on error
      return {
        userId,
        userEmail,
        shouldDeliver: true,
        channels: [NotificationChannel.Email, NotificationChannel.InApp],
        isDigest: false,
        digestFrequency: DigestFrequency.Immediate,
        reason: 'Using defaults due to error'
      };
    }
  }

  /**
   * Batch resolve delivery settings for multiple users
   */
  public async batchResolveDeliverySettings(
    recipients: Array<{ userId: number; userEmail: string }>,
    eventType: NotificationEventType,
    priority: Priority = Priority.Medium
  ): Promise<IResolvedDeliverySettings[]> {
    const results: IResolvedDeliverySettings[] = [];

    for (const recipient of recipients) {
      const settings = await this.resolveDeliverySettings(
        recipient.userId,
        recipient.userEmail,
        eventType,
        priority
      );
      results.push(settings);
    }

    return results;
  }

  // ============================================================================
  // DIGEST MANAGEMENT
  // ============================================================================

  /**
   * Queue a notification for digest delivery
   */
  public async queueForDigest(
    userId: number,
    eventType: NotificationEventType,
    subject: string,
    body: string,
    priority: Priority,
    digestFrequency: DigestFrequency,
    relatedItemId?: number,
    relatedItemType?: string
  ): Promise<number> {
    try {
      const scheduledTime = this.calculateNextDigestTime(digestFrequency);

      const result = await this.sp.web.lists.getByTitle(this.DIGEST_QUEUE_LIST)
        .items.add({
          UserId: userId,
          EventType: eventType,
          Subject: subject,
          Body: body,
          Priority: priority,
          RelatedItemId: relatedItemId,
          RelatedItemType: relatedItemType,
          ScheduledDigestTime: scheduledTime.toISOString(),
          CreatedDate: new Date().toISOString(),
          Status: 'Pending'
        });

      logger.info('NotificationPreferencesService',
        `Queued notification for digest delivery at ${scheduledTime.toISOString()}`);

      return result.data.Id;
    } catch (error) {
      logger.error('NotificationPreferencesService', 'Error queueing for digest', error);
      throw error;
    }
  }

  /**
   * Get pending digest items for a user
   */
  public async getPendingDigestItems(userId: number): Promise<IDigestItem[]> {
    try {
      const now = new Date().toISOString();

      const items = await this.sp.web.lists.getByTitle(this.DIGEST_QUEUE_LIST)
        .items
        .filter(`UserId eq ${userId} and Status eq 'Pending' and ScheduledDigestTime le '${now}'`)
        .orderBy('Priority', false)
        .orderBy('CreatedDate', true)();

      return items.map(item => ({
        id: item.Id,
        userId: item.UserId,
        eventType: item.EventType as NotificationEventType,
        subject: item.Subject,
        body: item.Body,
        priority: item.Priority as Priority,
        relatedItemId: item.RelatedItemId,
        relatedItemType: item.RelatedItemType,
        scheduledDigestTime: new Date(item.ScheduledDigestTime),
        createdDate: new Date(item.CreatedDate)
      }));
    } catch (error) {
      logger.error('NotificationPreferencesService', 'Error getting pending digest items', error);
      return [];
    }
  }

  /**
   * Mark digest items as sent
   */
  public async markDigestItemsSent(itemIds: number[]): Promise<void> {
    try {
      for (const itemId of itemIds) {
        await this.sp.web.lists.getByTitle(this.DIGEST_QUEUE_LIST)
          .items.getById(itemId)
          .update({
            Status: 'Sent',
            SentDate: new Date().toISOString()
          });
      }

      logger.info('NotificationPreferencesService',
        `Marked ${itemIds.length} digest items as sent`);
    } catch (error) {
      logger.error('NotificationPreferencesService', 'Error marking digest items sent', error);
    }
  }

  // ============================================================================
  // HELPER METHODS
  // ============================================================================

  /**
   * Get default event preferences
   */
  private getDefaultEventPreferences(): IEventChannelPreference[] {
    const defaults: IEventChannelPreference[] = [];

    // High priority events - immediate delivery
    const immediateEvents = [
      NotificationEventType.TaskAssigned,
      NotificationEventType.ApprovalRequired,
      NotificationEventType.ApprovalEscalated,
      NotificationEventType.WorkflowError,
      NotificationEventType.TaskOverdue
    ];

    // Normal events - can be digested
    const normalEvents = [
      NotificationEventType.TaskDueSoon,
      NotificationEventType.TaskCompleted,
      NotificationEventType.TaskUnblocked,
      NotificationEventType.ApprovalCompleted,
      NotificationEventType.ApprovalRejected,
      NotificationEventType.ProcessStarted,
      NotificationEventType.ProcessCompleted,
      NotificationEventType.ProcessBlocked,
      NotificationEventType.WorkflowStepComplete,
      NotificationEventType.SystemAnnouncement,
      NotificationEventType.Reminder
    ];

    for (const eventType of immediateEvents) {
      defaults.push({
        eventType,
        channels: [NotificationChannel.Email, NotificationChannel.InApp],
        enabled: true,
        digestFrequency: DigestFrequency.Immediate
      });
    }

    for (const eventType of normalEvents) {
      defaults.push({
        eventType,
        channels: [NotificationChannel.InApp],
        enabled: true,
        digestFrequency: DigestFrequency.Daily
      });
    }

    return defaults;
  }

  /**
   * Check if current time is within quiet hours
   */
  private isInQuietHours(quietHours: IQuietHours): boolean {
    const now = new Date();
    const currentDay = now.getDay();

    // Check if today is an excluded day
    if (quietHours.excludeDays.includes(currentDay)) {
      return true;  // Entire day is quiet
    }

    // Parse times
    const [startHour, startMin] = quietHours.startTime.split(':').map(Number);
    const [endHour, endMin] = quietHours.endTime.split(':').map(Number);

    const currentMinutes = now.getHours() * 60 + now.getMinutes();
    const startMinutes = startHour * 60 + startMin;
    const endMinutes = endHour * 60 + endMin;

    // Handle overnight quiet hours (e.g., 22:00 - 07:00)
    if (startMinutes > endMinutes) {
      return currentMinutes >= startMinutes || currentMinutes < endMinutes;
    }

    return currentMinutes >= startMinutes && currentMinutes < endMinutes;
  }

  /**
   * Calculate next digest delivery time based on frequency
   */
  private calculateNextDigestTime(frequency: DigestFrequency): Date {
    const now = new Date();

    switch (frequency) {
      case DigestFrequency.Hourly:
        now.setHours(now.getHours() + 1, 0, 0, 0);
        return now;

      case DigestFrequency.Daily:
        now.setDate(now.getDate() + 1);
        now.setHours(8, 0, 0, 0);  // 8 AM next day
        return now;

      case DigestFrequency.Weekly:
        const daysUntilMonday = (8 - now.getDay()) % 7 || 7;
        now.setDate(now.getDate() + daysUntilMonday);
        now.setHours(8, 0, 0, 0);  // 8 AM next Monday
        return now;

      default:
        return now;
    }
  }

  /**
   * Map SharePoint list item to preferences object
   */
  private mapItemToPreferences(item: Record<string, unknown>): IUserNotificationPreferences {
    return {
      id: item.Id as number,
      userId: item.UserId as number,
      userEmail: item.UserEmail as string,
      globalEnabled: item.GlobalEnabled as boolean,
      defaultChannels: JSON.parse(item.DefaultChannels as string || '["Email", "InApp"]'),
      defaultDigestFrequency: item.DefaultDigestFrequency as DigestFrequency || DigestFrequency.Immediate,
      eventPreferences: JSON.parse(item.EventPreferences as string || '[]'),
      quietHours: JSON.parse(item.QuietHours as string || '{"enabled": false}'),
      createdDate: new Date(item.CreatedDate as string),
      modifiedDate: new Date(item.ModifiedDate as string)
    };
  }

  /**
   * Update preferences cache
   */
  private updateCache(userId: number, preferences: IUserNotificationPreferences): void {
    const expiry = new Date();
    expiry.setMinutes(expiry.getMinutes() + this.CACHE_TTL_MINUTES);

    this.preferencesCache.set(userId, { preferences, expiry });
  }

  /**
   * Clear preferences cache for a user
   */
  public clearCache(userId?: number): void {
    if (userId) {
      this.preferencesCache.delete(userId);
    } else {
      this.preferencesCache.clear();
    }
  }
}

export default NotificationPreferencesService;
