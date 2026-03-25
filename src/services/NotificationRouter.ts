// @ts-nocheck
/**
 * NotificationRouter
 *
 * Central orchestrator for multi-channel notification delivery.
 * Routes notifications to Email, In-App, and Teams based on:
 * - Global admin settings (which channels are enabled per event)
 * - User preferences (per-user channel preferences)
 * - Quiet hours (Teams only)
 *
 * Usage:
 *   const router = new NotificationRouter(sp, context);
 *   await router.send({
 *     event: 'ack-required',
 *     recipientEmail: 'user@company.com',
 *     data: { policyId: 1, policyTitle: '...', deadline: '...' }
 *   });
 */

import { SPFI } from '@pnp/sp';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { TeamsNotificationService, ITeamsConfig, DEFAULT_TEAMS_CONFIG, TeamsCardType } from './TeamsNotificationService';

// ═══════════════════════════════════════════════════════════════
// TYPES
// ═══════════════════════════════════════════════════════════════

export type NotificationEvent =
  | 'policy-published' | 'policy-updated' | 'policy-retired'
  | 'ack-required' | 'ack-reminder-3day' | 'ack-reminder-1day' | 'ack-overdue' | 'ack-complete'
  | 'approval-request' | 'approval-approved' | 'approval-rejected' | 'approval-escalated' | 'approval-delegated'
  | 'quiz-assigned' | 'quiz-passed' | 'quiz-failed'
  | 'review-due' | 'review-overdue'
  | 'campaign-launched' | 'distribution-complete' | 'policy-assigned'
  | 'sla-breach' | 'violation-found' | 'policy-expiring'
  | 'weekly-digest' | 'welcome' | 'role-changed' | 'delegation-expiring';

export interface INotificationChannels {
  email: boolean;
  inApp: boolean;
  teams: boolean;
}

export interface INotificationEventConfig {
  event: NotificationEvent;
  channels: INotificationChannels;
  priority: 'low' | 'normal' | 'high' | 'urgent';
}

export interface ISendNotification {
  event: NotificationEvent;
  recipientEmail: string;
  recipientName?: string;
  data: Record<string, any>;
}

// Default channel configuration per event
export const DEFAULT_EVENT_CHANNELS: INotificationEventConfig[] = [
  // Acknowledgement
  { event: 'ack-required', channels: { email: true, inApp: true, teams: true }, priority: 'high' },
  { event: 'ack-reminder-3day', channels: { email: true, inApp: true, teams: true }, priority: 'normal' },
  { event: 'ack-reminder-1day', channels: { email: true, inApp: true, teams: true }, priority: 'high' },
  { event: 'ack-overdue', channels: { email: true, inApp: true, teams: true }, priority: 'urgent' },
  { event: 'ack-complete', channels: { email: false, inApp: true, teams: false }, priority: 'low' },
  // Approval
  { event: 'approval-request', channels: { email: true, inApp: true, teams: true }, priority: 'high' },
  { event: 'approval-approved', channels: { email: true, inApp: true, teams: true }, priority: 'normal' },
  { event: 'approval-rejected', channels: { email: true, inApp: true, teams: true }, priority: 'high' },
  { event: 'approval-escalated', channels: { email: true, inApp: true, teams: true }, priority: 'urgent' },
  { event: 'approval-delegated', channels: { email: true, inApp: true, teams: false }, priority: 'normal' },
  // Quiz
  { event: 'quiz-assigned', channels: { email: true, inApp: true, teams: true }, priority: 'normal' },
  { event: 'quiz-passed', channels: { email: false, inApp: true, teams: false }, priority: 'low' },
  { event: 'quiz-failed', channels: { email: true, inApp: true, teams: false }, priority: 'normal' },
  // Review
  { event: 'review-due', channels: { email: true, inApp: true, teams: false }, priority: 'normal' },
  { event: 'review-overdue', channels: { email: true, inApp: true, teams: true }, priority: 'high' },
  // Distribution
  { event: 'policy-published', channels: { email: true, inApp: true, teams: true }, priority: 'normal' },
  { event: 'policy-updated', channels: { email: true, inApp: true, teams: false }, priority: 'normal' },
  { event: 'policy-retired', channels: { email: true, inApp: true, teams: false }, priority: 'low' },
  { event: 'campaign-launched', channels: { email: true, inApp: true, teams: true }, priority: 'normal' },
  { event: 'distribution-complete', channels: { email: true, inApp: false, teams: false }, priority: 'low' },
  { event: 'policy-assigned', channels: { email: true, inApp: true, teams: true }, priority: 'normal' },
  // Compliance
  { event: 'sla-breach', channels: { email: true, inApp: true, teams: true }, priority: 'urgent' },
  { event: 'violation-found', channels: { email: true, inApp: true, teams: true }, priority: 'urgent' },
  { event: 'policy-expiring', channels: { email: true, inApp: true, teams: false }, priority: 'normal' },
  // System
  { event: 'weekly-digest', channels: { email: true, inApp: false, teams: true }, priority: 'low' },
  { event: 'welcome', channels: { email: true, inApp: true, teams: true }, priority: 'normal' },
  { event: 'role-changed', channels: { email: true, inApp: true, teams: false }, priority: 'normal' },
  { event: 'delegation-expiring', channels: { email: true, inApp: true, teams: false }, priority: 'normal' },
];

// Map notification events to Teams card types
const EVENT_TO_CARD_TYPE: Partial<Record<NotificationEvent, TeamsCardType>> = {
  'policy-published': 'policy-published',
  'ack-required': 'ack-required',
  'ack-reminder-3day': 'ack-reminder',
  'ack-reminder-1day': 'ack-reminder',
  'ack-overdue': 'ack-reminder',
  'approval-request': 'approval-request',
  'approval-approved': 'approval-result',
  'approval-rejected': 'approval-result',
  'quiz-assigned': 'quiz-assigned',
  'sla-breach': 'sla-breach',
  'weekly-digest': 'weekly-digest',
};

// ═══════════════════════════════════════════════════════════════
// ROUTER
// ═══════════════════════════════════════════════════════════════

export class NotificationRouter {
  private sp: SPFI;
  private context: WebPartContext;
  private teamsService: TeamsNotificationService | null = null;
  private eventConfigs: INotificationEventConfig[];
  private siteUrl: string;

  constructor(sp: SPFI, context: WebPartContext) {
    this.sp = sp;
    this.context = context;
    this.eventConfigs = [...DEFAULT_EVENT_CHANNELS];
    this.siteUrl = context.pageContext?.web?.absoluteUrl || '';
  }

  /**
   * Initialize the router — loads config from SP and creates Teams service
   */
  public async initialize(): Promise<void> {
    try {
      // Load Teams config
      const configItems = await this.sp.web.lists.getByTitle('PM_Configuration')
        .items.filter("Category eq 'Notifications' or Category eq 'Teams'")
        .select('ConfigKey', 'ConfigValue')
        .top(50)();

      const configMap: Record<string, string> = {};
      configItems.forEach((item: any) => { configMap[item.ConfigKey] = item.ConfigValue; });

      // Parse Teams config
      const teamsConfig: ITeamsConfig = {
        ...DEFAULT_TEAMS_CONFIG,
        enabled: configMap['Notifications.Teams.Enabled'] === 'true',
        channelWebhookUrl: configMap['Notifications.Teams.WebhookUrl'] || '',
        enableActivityFeed: configMap['Notifications.Teams.ActivityFeed'] !== 'false',
        enableAdaptiveCards: configMap['Notifications.Teams.AdaptiveCards'] !== 'false',
        enableChannelPosts: configMap['Notifications.Teams.ChannelPosts'] === 'true',
        quietHoursStart: parseInt(configMap['Notifications.Teams.QuietStart'] || '20', 10),
        quietHoursEnd: parseInt(configMap['Notifications.Teams.QuietEnd'] || '7', 10),
        respectQuietHours: configMap['Notifications.Teams.QuietHours'] !== 'false'
      };

      this.teamsService = new TeamsNotificationService(this.context, teamsConfig);

      // Parse per-event channel overrides
      const overridesJson = configMap['Notifications.EventChannels'];
      if (overridesJson) {
        try {
          const overrides = JSON.parse(overridesJson);
          this.eventConfigs = this.eventConfigs.map(ec => {
            const override = overrides.find((o: any) => o.event === ec.event);
            return override ? { ...ec, channels: { ...ec.channels, ...override.channels } } : ec;
          });
        } catch { /* use defaults */ }
      }
    } catch (err) {
      console.error('[NotificationRouter] initialize failed:', err);
    }
  }

  /**
   * Send a notification across all configured channels
   */
  public async send(notification: ISendNotification): Promise<{ email: boolean; inApp: boolean; teams: boolean }> {
    const result = { email: false, inApp: false, teams: false };
    const eventConfig = this.eventConfigs.find(ec => ec.event === notification.event);
    if (!eventConfig) return result;

    const channels = eventConfig.channels;

    // Email — queue to PM_EmailQueue
    if (channels.email) {
      try {
        await this.sp.web.lists.getByTitle('PM_NotificationQueue').items.add({
          Title: `${notification.event}: ${notification.data.policyTitle || 'Notification'}`,
          RecipientEmail: notification.recipientEmail,
          RecipientName: notification.recipientName || '',
          PolicyTitle: notification.data.policyTitle || '',
          PolicyId: notification.data.policyId || 0,
          Message: notification.data.body || notification.data.subject || '',
          Priority: eventConfig.priority === 'urgent' ? 'High' : 'Normal',
          Status: 'Pending',
          NotificationType: notification.event,
          Channel: 'Email'
        });
        result.email = true;
      } catch (err) {
        console.error('[NotificationRouter] email queue failed:', err);
      }
    }

    // In-App — write to PM_Notifications
    if (channels.inApp) {
      try {
        await this.sp.web.lists.getByTitle('PM_Notifications').items.add({
          Title: notification.data.policyTitle || notification.event,
          Message: notification.data.message || notification.data.summary || '',
          Type: 'Policy',
          Priority: eventConfig.priority === 'urgent' ? 'High' : eventConfig.priority === 'high' ? 'High' : 'Normal',
          IsRead: false,
          RelatedItemType: 'Policy',
          RelatedItemId: notification.data.policyId || 0
        });
        result.inApp = true;
      } catch (err) {
        console.error('[NotificationRouter] in-app notification failed:', err);
      }
    }

    // Teams — send via TeamsNotificationService
    if (channels.teams && this.teamsService?.isAvailable()) {
      try {
        const cardType = EVENT_TO_CARD_TYPE[notification.event];
        if (cardType) {
          result.teams = await this.teamsService.sendAdaptiveCard({
            recipientEmail: notification.recipientEmail,
            cardType,
            data: notification.data
          });
        } else {
          // Fallback to activity feed for events without card templates
          result.teams = await this.teamsService.sendActivityFeed(
            notification.recipientEmail,
            notification.data.policyTitle || 'Policy Manager',
            notification.data.message || notification.event,
            notification.data.policyId
              ? `${this.siteUrl}/SitePages/PolicyDetails.aspx?policyId=${notification.data.policyId}`
              : undefined
          );
        }

        // Also post to channel if it's a broadcast event
        if (['policy-published', 'campaign-launched', 'sla-breach'].includes(notification.event)) {
          await this.teamsService.sendChannelCard(
            cardType || 'policy-published',
            notification.data
          );
        }
      } catch (err) {
        console.error('[NotificationRouter] Teams notification failed:', err);
      }
    }

    return result;
  }

  /**
   * Get the current event channel configurations
   */
  public getEventConfigs(): INotificationEventConfig[] {
    return [...this.eventConfigs];
  }

  /**
   * Update event channel configurations
   */
  public updateEventConfigs(configs: INotificationEventConfig[]): void {
    this.eventConfigs = configs;
  }

  /**
   * Save event channel configurations to SP
   */
  public async saveEventConfigs(): Promise<void> {
    const json = JSON.stringify(this.eventConfigs.map(ec => ({
      event: ec.event,
      channels: ec.channels,
      priority: ec.priority
    })));

    try {
      const list = this.sp.web.lists.getByTitle('PM_Configuration');
      const items = await list.items.filter("ConfigKey eq 'Notifications.EventChannels'").top(1)();
      if (items.length > 0) {
        await list.items.getById(items[0].Id).update({ ConfigValue: json });
      } else {
        await list.items.add({
          Title: 'Notification Event Channels',
          ConfigKey: 'Notifications.EventChannels',
          ConfigValue: json,
          Category: 'Notifications',
          IsActive: true
        });
      }
    } catch (err) {
      console.error('[NotificationRouter] saveEventConfigs failed:', err);
      throw err;
    }
  }
}
