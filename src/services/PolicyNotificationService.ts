/**
 * Policy Notification Service
 * Handles email notifications, reminders, and alerts for Policy Management
 */

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/site-users/web';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import { IPolicy, IPolicyAcknowledgement, AcknowledgementStatus } from '../models/IPolicy';
import { logger } from './LoggingService';
import { NotificationLists } from '../constants/SharePointListNames';
import { DwxNotificationService, DwxNotificationType, DwxNotificationPriority, DwxNotificationCategory } from '@dwx/core';
import { escapeHtml } from '../utils/sanitizeHtml';
import { ValidationUtils } from '../utils/ValidationUtils';
import { NotificationRouter } from './NotificationRouter';

/**
 * Policy notification types
 */
export type PolicyNotificationType =
  | 'NewPolicy'
  | 'PolicyUpdated'
  | 'AcknowledgementRequired'
  | 'ApprovalRequired'
  | 'Reminder3Day'
  | 'Reminder1Day'
  | 'Overdue'
  | 'AcknowledgementComplete'
  | 'PolicyExpiring'
  | 'PolicyApproved'
  | 'PolicyRejected'
  | 'DelegationRequest';

/**
 * Notification configuration
 */
export interface IPolicyNotification {
  recipientId: number;
  recipientEmail?: string;
  recipientName?: string;
  notificationType: PolicyNotificationType;
  subject: string;
  body: string;
  policyId?: number;
  sendEmail: boolean;
  sendInApp: boolean;
}

/**
 * Reminder schedule configuration
 */
export interface IReminderSchedule {
  policyId: number;
  userId: number;
  dueDate: Date;
  reminder3DaySent: boolean;
  reminder1DaySent: boolean;
  overdueSent: boolean;
}

/**
 * Policy Notification Service
 */
export class PolicyNotificationService {
  private sp: SPFI;
  private siteUrl: string;
  private dwxNotifications: DwxNotificationService | null;
  public notificationRouter: NotificationRouter | null = null;
  // private readonly NOTIFICATION_LIST = NotificationLists.POLICY_NOTIFICATIONS; // Now using 'PM_Notifications' directly
  private readonly REMINDER_SCHEDULE_LIST = NotificationLists.REMINDER_SCHEDULE;

  constructor(sp: SPFI, siteUrl: string, dwxNotificationService?: DwxNotificationService) {
    this.sp = sp;
    this.siteUrl = siteUrl;
    this.dwxNotifications = dwxNotificationService || null;
  }

  /**
   * Set the NotificationRouter for multi-channel delivery (Teams Adaptive Cards, etc.)
   */
  public setNotificationRouter(router: NotificationRouter): void {
    this.notificationRouter = router;
  }

  // ============================================================================
  // TEMPLATE ENGINE — loads admin-configured email templates from PM_EmailTemplates
  // ============================================================================

  private templateCache: Map<string, { subject: string; body: string; isActive: boolean }> = new Map();
  private templateCacheTime = 0;

  /**
   * Load email template by event name from PM_EmailTemplates.
   * Caches for 5 minutes to avoid repeated SP queries.
   */
  private async loadEmailTemplate(eventName: string): Promise<{ subject: string; body: string } | null> {
    // Check cache (5 min TTL)
    const now = Date.now();
    if (now - this.templateCacheTime > 300000) {
      this.templateCache.clear();
      this.templateCacheTime = now;
      try {
        const items = await this.sp.web.lists.getByTitle('PM_EmailTemplates')
          .items.filter("IsActive eq 1")
          .select('Event', 'Subject', 'Body', 'IsActive')
          .top(50)();
        for (const item of items) {
          this.templateCache.set(item.Event, { subject: item.Subject, body: item.Body, isActive: item.IsActive });
        }
      } catch {
        // PM_EmailTemplates may not exist — fall back to hardcoded
        return null;
      }
    }
    const template = this.templateCache.get(eventName);
    if (template && template.isActive) {
      return { subject: template.subject, body: template.body };
    }
    return null;
  }

  /**
   * Replace merge tags in template string.
   * Tags use {{TagName}} format.
   */
  private replaceMergeTags(template: string, data: Record<string, string>): string {
    let result = template;
    for (const [key, value] of Object.entries(data)) {
      result = result.replace(new RegExp(`\\{\\{${key}\\}\\}`, 'g'), value || '');
    }
    return result;
  }

  /**
   * Queue a templated email to PM_NotificationQueue.
   * Loads admin template first; falls back to provided fallback subject/body.
   */
  public async queueTemplatedEmail(opts: {
    eventName: string;
    recipientEmail: string;
    recipientName: string;
    mergeData: Record<string, string>;
    fallbackSubject: string;
    fallbackBody: string;
    policyId?: number;
    priority?: string;
  }): Promise<void> {
    try {
      // Try admin template first
      const template = await this.loadEmailTemplate(opts.eventName);
      let subject = opts.fallbackSubject;
      let body = opts.fallbackBody;

      if (template) {
        subject = this.replaceMergeTags(template.subject, opts.mergeData);
        body = this.replaceMergeTags(template.body, opts.mergeData);
      }

      // Wrap body in email shell
      const siteUrl = this.sp.web.toUrl().replace('/_api/web', '');
      const isReviewEvent = ['review-required', 'approval-request', 'review-withdrawn'].includes(opts.eventName);
      const policyUrl = opts.mergeData.PolicyUrl || `${siteUrl}/SitePages/PolicyDetails.aspx?policyId=${opts.policyId || 0}${isReviewEvent ? '&mode=review' : ''}`;
      const htmlBody = this.buildEmailShell({
        headerGradient: 'linear-gradient(135deg, #0d9488 0%, #0f766e 100%)',
        headerIcon: '&#x1F4CB;',
        headerTitle: subject,
        content: body,
        footerText: 'First Digital — DWx Policy Manager',
        ctaUrl: policyUrl,
        ctaLabel: 'View in Policy Manager',
        ctaColor: '#0d9488'
      });

      // Queue to PM_NotificationQueue (two-step write to guarantee QueueStatus)
      const qResult = await this.sp.web.lists.getByTitle('PM_NotificationQueue').items.add({
        Title: subject,
        RecipientEmail: opts.recipientEmail,
        RecipientName: opts.recipientName,
        PolicyId: opts.policyId || 0,
        PolicyTitle: opts.mergeData.PolicyTitle || '',
        NotificationType: opts.eventName,
        Channel: 'Email',
        Message: htmlBody,
        QueueStatus: 'Pending',
        Priority: opts.priority || 'Normal'
      });
      try { const qId = qResult?.data?.Id; if (qId) await this.sp.web.lists.getByTitle('PM_NotificationQueue').items.getById(qId).update({ QueueStatus: 'Pending' }); } catch { /* */ }
    } catch (err) {
      logger.warn('PolicyNotificationService', `Failed to queue templated email for ${opts.eventName}:`, err);
    }
  }

  // ============================================================================
  // PUBLIC NOTIFICATION METHODS
  // ============================================================================

  /**
   * Send notification for new policy published
   */
  public async sendNewPolicyNotification(
    policy: IPolicy,
    recipientIds: number[]
  ): Promise<void> {
    const emailBody = this.buildNewPolicyEmail(policy);
    let failCount = 0;
    for (const recipientId of recipientIds) {
      try {
        await this.sendNotification({
          recipientId,
          notificationType: 'NewPolicy',
          subject: `New Policy Published: ${policy.PolicyName}`,
          body: emailBody,
          policyId: policy.Id,
          sendEmail: true,
          sendInApp: true
        });
      } catch (err) {
        failCount++;
        logger.warn('PolicyNotificationService', `Failed to notify recipient ${recipientId}:`, err);
      }
    }
    if (failCount > 0) {
      logger.error('PolicyNotificationService',
        `Failed to send new policy notification to ${failCount}/${recipientIds.length} recipients`);
    }
  }

  /**
   * Send notification for policy update
   */
  public async sendPolicyUpdateNotification(
    policy: IPolicy,
    recipientIds: number[],
    changeDescription: string
  ): Promise<void> {
    const emailBody = this.buildPolicyUpdateEmail(policy, changeDescription);
    let failCount = 0;
    for (const recipientId of recipientIds) {
      try {
        await this.sendNotification({
          recipientId,
          notificationType: 'PolicyUpdated',
          subject: `Policy Updated: ${policy.PolicyName}`,
          body: emailBody,
          policyId: policy.Id,
          sendEmail: true,
          sendInApp: true
        });
      } catch (err) {
        failCount++;
        logger.warn('PolicyNotificationService', `Failed to notify recipient ${recipientId}:`, err);
      }
    }
    if (failCount > 0) {
      logger.error('PolicyNotificationService',
        `Failed to send policy update notification to ${failCount}/${recipientIds.length} recipients`);
    }
  }

  /**
   * Send acknowledgement required notification
   */
  public async sendAcknowledgementRequiredNotification(
    policy: IPolicy,
    acknowledgement: IPolicyAcknowledgement
  ): Promise<void> {
    try {
      const notification: IPolicyNotification = {
        recipientId: acknowledgement.AckUserId,
        notificationType: 'AcknowledgementRequired',
        subject: `Action Required: Acknowledge Policy - ${policy.PolicyName}`,
        body: this.buildAcknowledgementRequiredEmail(policy, acknowledgement),
        policyId: policy.Id,
        sendEmail: true,
        sendInApp: true
      };

      await this.sendNotification(notification);

      // Schedule reminders if due date is set
      if (acknowledgement.DueDate) {
        await this.scheduleReminders(policy.Id!, acknowledgement.AckUserId, new Date(acknowledgement.DueDate));
      }
    } catch (error) {
      logger.error('PolicyNotificationService', 'Failed to send acknowledgement required notification:', error);
    }
  }

  /**
   * Send 3-day reminder notification
   */
  public async sendReminder3DayNotification(
    policy: IPolicy,
    acknowledgement: IPolicyAcknowledgement
  ): Promise<void> {
    try {
      const notification: IPolicyNotification = {
        recipientId: acknowledgement.AckUserId,
        notificationType: 'Reminder3Day',
        subject: `Reminder: 3 Days to Acknowledge - ${policy.PolicyName}`,
        body: this.buildReminder3DayEmail(policy, acknowledgement),
        policyId: policy.Id,
        sendEmail: true,
        sendInApp: true
      };

      await this.sendNotification(notification);
    } catch (error) {
      logger.error('PolicyNotificationService', 'Failed to send 3-day reminder:', error);
    }
  }

  /**
   * Send 1-day reminder notification
   */
  public async sendReminder1DayNotification(
    policy: IPolicy,
    acknowledgement: IPolicyAcknowledgement
  ): Promise<void> {
    try {
      const notification: IPolicyNotification = {
        recipientId: acknowledgement.AckUserId,
        notificationType: 'Reminder1Day',
        subject: `Urgent Reminder: Acknowledge Tomorrow - ${policy.PolicyName}`,
        body: this.buildReminder1DayEmail(policy, acknowledgement),
        policyId: policy.Id,
        sendEmail: true,
        sendInApp: true
      };

      await this.sendNotification(notification);
    } catch (error) {
      logger.error('PolicyNotificationService', 'Failed to send 1-day reminder:', error);
    }
  }

  /**
   * Send overdue notification
   */
  public async sendOverdueNotification(
    policy: IPolicy,
    acknowledgement: IPolicyAcknowledgement,
    daysOverdue: number
  ): Promise<void> {
    try {
      const notification: IPolicyNotification = {
        recipientId: acknowledgement.AckUserId,
        notificationType: 'Overdue',
        subject: `OVERDUE: Policy Acknowledgement Required - ${policy.PolicyName}`,
        body: this.buildOverdueEmail(policy, acknowledgement, daysOverdue),
        policyId: policy.Id,
        sendEmail: true,
        sendInApp: true
      };

      await this.sendNotification(notification);

      // Also notify the user's manager if available
      if ((acknowledgement as any).ManagerId) {
        await this.sendManagerOverdueAlert(policy, acknowledgement, daysOverdue);
      }
    } catch (error) {
      logger.error('PolicyNotificationService', 'Failed to send overdue notification:', error);
    }
  }

  /**
   * Send acknowledgement complete notification
   */
  public async sendAcknowledgementCompleteNotification(
    policy: IPolicy,
    acknowledgement: IPolicyAcknowledgement
  ): Promise<void> {
    try {
      const notification: IPolicyNotification = {
        recipientId: acknowledgement.AckUserId,
        notificationType: 'AcknowledgementComplete',
        subject: `Confirmed: Policy Acknowledged - ${policy.PolicyName}`,
        body: this.buildAcknowledgementCompleteEmail(policy, acknowledgement),
        policyId: policy.Id,
        sendEmail: true,
        sendInApp: true
      };

      await this.sendNotification(notification);
    } catch (error) {
      logger.error('PolicyNotificationService', 'Failed to send acknowledgement complete notification:', error);
    }
  }

  /**
   * Send policy expiring notification to admins
   */
  public async sendPolicyExpiringNotification(
    policy: IPolicy,
    adminIds: number[],
    daysUntilExpiry: number
  ): Promise<void> {
    try {
      for (const adminId of adminIds) {
        const notification: IPolicyNotification = {
          recipientId: adminId,
          notificationType: 'PolicyExpiring',
          subject: `Policy Expiring in ${daysUntilExpiry} Days: ${policy.PolicyName}`,
          body: this.buildPolicyExpiringEmail(policy, daysUntilExpiry),
          policyId: policy.Id,
          sendEmail: true,
          sendInApp: true
        };

        await this.sendNotification(notification);
      }
    } catch (error) {
      logger.error('PolicyNotificationService', 'Failed to send policy expiring notification:', error);
    }
  }

  /**
   * Send policy approval notification
   */
  public async sendPolicyApprovalNotification(
    policy: IPolicy,
    authorId: number,
    approverName: string
  ): Promise<void> {
    try {
      const notification: IPolicyNotification = {
        recipientId: authorId,
        notificationType: 'PolicyApproved',
        subject: `Policy Approved: ${policy.PolicyName}`,
        body: this.buildPolicyApprovalEmail(policy, approverName),
        policyId: policy.Id,
        sendEmail: true,
        sendInApp: true
      };

      await this.sendNotification(notification);
    } catch (error) {
      logger.error('PolicyNotificationService', 'Failed to send policy approval notification:', error);
    }
  }

  /**
   * Send policy rejection notification
   */
  public async sendPolicyRejectionNotification(
    policy: IPolicy,
    authorId: number,
    approverName: string,
    reason: string
  ): Promise<void> {
    try {
      const notification: IPolicyNotification = {
        recipientId: authorId,
        notificationType: 'PolicyRejected',
        subject: `Policy Requires Revision: ${policy.PolicyName}`,
        body: this.buildPolicyRejectionEmail(policy, approverName, reason),
        policyId: policy.Id,
        sendEmail: true,
        sendInApp: true
      };

      await this.sendNotification(notification);
    } catch (error) {
      logger.error('PolicyNotificationService', 'Failed to send policy rejection notification:', error);
    }
  }

  /**
   * Send notification to reviewers when a policy is submitted for review
   */
  public async sendSubmittedForReviewNotification(
    policy: IPolicy,
    reviewerIds: number[],
    submitterName: string
  ): Promise<void> {
    const siteUrl = this.sp.web.toUrl().replace('/_api/web', '');
    const policyUrl = `${siteUrl}/SitePages/PolicyDetails.aspx?policyId=${policy.Id}&mode=review`;

    // Send email + in-app notification to each reviewer
    let failCount = 0;
    for (const reviewerId of reviewerIds) {
      try {
        // Resolve reviewer email/name
        const reviewer = await this.sp.web.siteUsers.getById(reviewerId).select('Email', 'Title')();
        if (!reviewer?.Email) continue;

        // Queue templated email via PM_NotificationQueue
        await this.queueTemplatedEmail({
          eventName: 'review-required',
          recipientEmail: reviewer.Email,
          recipientName: reviewer.Title || '',
          policyId: policy.Id,
          priority: 'High',
          mergeData: {
            PolicyTitle: policy.PolicyName,
            PolicyNumber: policy.PolicyNumber || '',
            AuthorName: submitterName,
            RecipientName: reviewer.Title || '',
            Category: policy.PolicyCategory || '',
            RiskLevel: policy.ComplianceRisk || 'Medium',
            PolicyUrl: policyUrl
          },
          fallbackSubject: `Review Required: ${policy.PolicyName}`,
          fallbackBody: `<p>${submitterName} has submitted <strong>${policy.PolicyName}</strong> for your review.</p><p><a href="${policyUrl}">Review Policy</a></p>`
        });

        // In-app notification
        try {
          await this.sp.web.lists.getByTitle('PM_Notifications').items.add({
            Title: `Review Required: ${policy.PolicyName}`,
            RecipientId: reviewerId,
            Message: `${submitterName} has submitted "${policy.PolicyName}" for your review.`,
            Type: 'Policy',
            RelatedItemId: policy.Id,
            ActionUrl: policyUrl,
            Priority: 'High',
            IsRead: false
          });
        } catch { /* in-app notification is best-effort */ }
      } catch (err) {
        failCount++;
        logger.warn('PolicyNotificationService', `Failed to notify reviewer ${reviewerId}:`, err);
      }
    }
    if (failCount > 0) {
      logger.error('PolicyNotificationService',
        `Failed to send review notification to ${failCount}/${reviewerIds.length} recipients`);
    }
  }

  // ============================================================================
  // REMINDER SCHEDULER
  // ============================================================================

  /**
   * Schedule reminders for a policy acknowledgement
   */
  public async scheduleReminders(
    policyId: number,
    userId: number,
    dueDate: Date
  ): Promise<void> {
    try {
      // Check if schedule already exists
      const existing = await this.sp.web.lists
        .getByTitle(this.REMINDER_SCHEDULE_LIST)
        .items.filter(`PolicyId eq ${ValidationUtils.validateInteger(policyId, 'policyId', 1)} and UserId eq ${ValidationUtils.validateInteger(userId, 'userId', 1)}`)
        .top(1)();

      if (existing.length > 0) {
        // Update existing schedule
        await this.sp.web.lists
          .getByTitle(this.REMINDER_SCHEDULE_LIST)
          .items.getById(existing[0].Id)
          .update({
            DueDate: dueDate.toISOString(),
            Reminder3DaySent: false,
            Reminder1DaySent: false,
            OverdueSent: false
          });
      } else {
        // Create new schedule
        await this.sp.web.lists
          .getByTitle(this.REMINDER_SCHEDULE_LIST)
          .items.add({
            Title: `Policy ${policyId} - User ${userId}`,
            PolicyId: policyId,
            UserId: userId,
            DueDate: dueDate.toISOString(),
            Reminder3DaySent: false,
            Reminder1DaySent: false,
            OverdueSent: false
          });
      }

      logger.info('PolicyNotificationService', `Scheduled reminders for policy ${policyId}, user ${userId}`);
    } catch (error) {
      logger.error('PolicyNotificationService', 'Failed to schedule reminders:', error);
    }
  }

  /**
   * Process pending reminders - should be called by a timer job
   */
  public async processReminders(): Promise<{
    processed: number;
    reminders3Day: number;
    reminders1Day: number;
    overdueAlerts: number;
  }> {
    const stats = { processed: 0, reminders3Day: 0, reminders1Day: 0, overdueAlerts: 0 };

    try {
      const now = new Date();

      // Load SLA targets from admin configuration (PM_Configuration)
      // Acknowledgement SLA drives reminder thresholds
      let warningDays = 3;  // default: 3 days before deadline
      let urgentDays = 1;   // default: 1 day before deadline
      try {
        const slaItems = await this.sp.web.lists
          .getByTitle('PM_Configuration')
          .items.filter("substringof('Admin.SLA', ConfigKey)")
          .select('ConfigKey', 'ConfigValue')
          .top(10)() as any[];

        // Look for acknowledgement SLA config
        for (const item of slaItems) {
          if (item.ConfigKey?.includes('Acknowledgement') && item.ConfigValue) {
            try {
              const sla = JSON.parse(item.ConfigValue);
              if (sla.WarningThresholdDays) warningDays = Number(sla.WarningThresholdDays);
              if (sla.UrgentThresholdDays) urgentDays = Number(sla.UrgentThresholdDays);
            } catch { /* use defaults */ }
          }
        }
      } catch {
        // PM_Configuration may not exist — use hardcoded defaults
      }

      // Get all pending schedules
      const schedules = await this.sp.web.lists
        .getByTitle(this.REMINDER_SCHEDULE_LIST)
        .items.filter(`(Reminder3DaySent eq false or Reminder1DaySent eq false or OverdueSent eq false)`)
        .top(500)() as any[];

      if (schedules.length === 0) return stats;

      // Batch-load policies and acknowledgements to avoid N+1 queries
      const policyIdSet = new Set<number>();
      schedules.forEach((s: any) => policyIdSet.add(s.PolicyId));
      const uniquePolicyIds = Array.from(policyIdSet);
      const policyMap = new Map<number, any>();
      const ackMap = new Map<string, any>();

      // Load policies in batches of 50
      for (let i = 0; i < uniquePolicyIds.length; i += 50) {
        const batch = uniquePolicyIds.slice(i, i + 50);
        try {
          const filter = batch.map(function(id) { return 'Id eq ' + Number(id); }).join(' or ');
          const policies = await this.sp.web.lists
            .getByTitle('PM_Policies')
            .items.filter(filter)
            .select('Id', 'Title', 'PolicyName', 'PolicyNumber', 'PolicyCategory', 'ComplianceRisk')
            .top(50)() as any[];
          policies.forEach(function(p: any) { policyMap.set(p.Id, p); });
        } catch { /* continue with partial data */ }
      }

      // Load acknowledgements in batches
      for (let i = 0; i < schedules.length; i += 50) {
        const batch = schedules.slice(i, i + 50);
        try {
          const filter = batch.map(function(s: any) { return '(PolicyId eq ' + Number(s.PolicyId) + ' and UserId eq ' + Number(s.UserId) + ')'; }).join(' or ');
          const acks = await this.sp.web.lists
            .getByTitle('PM_PolicyAcknowledgements')
            .items.filter(filter)
            .select('Id', 'PolicyId', 'UserId', 'AckStatus', 'UserEmail', 'UserName')
            .top(50)() as any[];
          acks.forEach(function(a: any) { ackMap.set(a.PolicyId + '_' + a.UserId, a); });
        } catch { /* continue with partial data */ }
      }

      for (const schedule of schedules) {
        try {
          stats.processed++;
          const dueDate = new Date(schedule.DueDate);
          const daysToDue = Math.ceil((dueDate.getTime() - now.getTime()) / (1000 * 60 * 60 * 24));

          const policy = policyMap.get(schedule.PolicyId);
          if (!policy) continue;

          const acknowledgement = ackMap.get(schedule.PolicyId + '_' + schedule.UserId);
          if (!acknowledgement || acknowledgement.AckStatus === AcknowledgementStatus.Acknowledged) {
            await this.sp.web.lists
              .getByTitle(this.REMINDER_SCHEDULE_LIST)
              .items.getById(schedule.Id)
              .delete();
            continue;
          }

          // Warning reminder (SLA-driven, default 3 days)
          if (daysToDue <= warningDays && daysToDue > urgentDays && !schedule.Reminder3DaySent) {
            await this.sendReminder3DayNotification(policy, acknowledgement);
            await this.updateReminderSchedule(schedule.Id, { reminder3DaySent: true });
            stats.reminders3Day++;
          }

          // Urgent reminder (SLA-driven, default 1 day)
          if (daysToDue <= urgentDays && daysToDue >= 0 && !schedule.Reminder1DaySent) {
            await this.sendReminder1DayNotification(policy, acknowledgement);
            await this.updateReminderSchedule(schedule.Id, { reminder1DaySent: true });
            stats.reminders1Day++;
          }

          // Overdue
          if (daysToDue < 0 && !schedule.OverdueSent) {
            await this.sendOverdueNotification(policy, acknowledgement, Math.abs(daysToDue));
            await this.updateReminderSchedule(schedule.Id, { overdueSent: true });
            stats.overdueAlerts++;
          }
        } catch (scheduleErr) {
          logger.warn('PolicyNotificationService', 'Failed processing schedule ' + schedule.Id + ':', scheduleErr);
        }
      }

      logger.info('PolicyNotificationService', `Processed ${stats.processed} reminder schedules`, stats);
      return stats;
    } catch (error) {
      logger.error('PolicyNotificationService', 'Failed to process reminders:', error);
      return stats;
    }
  }

  // ============================================================================
  // EMAIL TEMPLATE BUILDERS — Forest Teal Brand Design System
  // ============================================================================

  /**
   * Wraps email content in a branded, responsive HTML shell.
   * @param headerGradient CSS gradient for the header bar
   * @param headerIcon Emoji icon for the header
   * @param headerTitle Title text for the header
   * @param bodyBg Background color for the content area
   * @param content Inner HTML content
   * @param footerText Footer disclaimer text
   */
  private buildEmailShell(opts: {
    headerGradient: string; headerIcon: string; headerTitle: string;
    bodyBg?: string; content: string; footerText: string;
    ctaUrl: string; ctaLabel: string; ctaColor: string;
  }): string {
    const { headerGradient, headerTitle, content, ctaUrl, ctaLabel, ctaColor } = opts;
    // Extract gradient colours from "linear-gradient(135deg, #xxx, #yyy)" format
    const gradientMatch = headerGradient.match(/#[0-9a-fA-F]{6}/g) || ['#0d9488', '#0f766e'];
    const gradStart = gradientMatch[0] || '#0d9488';
    const gradEnd = gradientMatch[1] || gradientMatch[0] || '#0f766e';
    const F = "'Segoe UI', -apple-system, BlinkMacSystemFont, Roboto, Helvetica, Arial, sans-serif";

    return `<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="background-color:#f1f5f9;">
  <tr>
    <td align="center" style="padding:32px 16px;">
      <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="600" style="max-width:600px; width:100%; border-radius:12px; overflow:hidden; box-shadow:0 4px 24px rgba(0,0,0,0.08);">
        <tr>
          <td style="background:linear-gradient(135deg, ${gradStart} 0%, ${gradEnd} 100%); padding:0;">
            <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%">
              <tr>
                <td style="padding:20px 40px 18px 40px;">
                  <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%">
                    <tr>
                      <td valign="middle" style="font-family:${F};">
                        <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="padding-bottom:10px;">
                          <tr>
                            <td style="font-size:11px; font-weight:600; letter-spacing:1.5px; text-transform:uppercase; color:rgba(255,255,255,0.6);">First Digital &bull; DWx Policy Manager</td>
                          </tr>
                        </table>
                        <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%">
                          <tr>
                            <td style="font-size:20px; font-weight:700; color:#ffffff; line-height:1.3; letter-spacing:-0.3px;">${headerTitle}</td>
                            <td width="44" valign="middle" align="right">
                              <table role="presentation" cellpadding="0" cellspacing="0" border="0">
                                <tr><td style="width:44px; height:44px; border-radius:50%; background-color:rgba(255,255,255,0.1); font-size:1px; line-height:1px;">&nbsp;</td></tr>
                              </table>
                            </td>
                          </tr>
                        </table>
                      </td>
                    </tr>
                  </table>
                </td>
              </tr>
            </table>
          </td>
        </tr>
        <tr>
          <td style="background-color:#ffffff; padding:28px 40px 24px 40px;">
            ${content}
            <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%">
              <tr><td style="height:28px; font-size:1px; line-height:1px;">&nbsp;</td></tr>
            </table>
            <table role="presentation" cellpadding="0" cellspacing="0" border="0" align="center">
              <tr>
                <td style="background-color:${ctaColor}; border-radius:8px;">
                  <a href="${ctaUrl}" target="_blank" style="display:inline-block; padding:14px 48px; font-family:${F}; font-size:14px; font-weight:600; color:#ffffff; text-decoration:none; letter-spacing:0.3px;">${ctaLabel}</a>
                </td>
              </tr>
            </table>
            <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%">
              <tr>
                <td align="center" style="padding-top:12px; font-family:${F}; font-size:11px; color:#94a3b8;">
                  Or copy this link: <a href="${ctaUrl}" style="color:#64748b; text-decoration:underline; word-break:break-all;">${ctaUrl}</a>
                </td>
              </tr>
            </table>
          </td>
        </tr>
        <tr>
          <td style="background-color:#f8fafc; border-top:1px solid #e2e8f0; padding:20px 40px;">
            <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%">
              <tr>
                <td style="font-family:${F}; font-size:11px; color:#94a3b8; line-height:1.6;">First Digital &mdash; DWx Policy Manager<br><span style="color:#cbd5e1;">Policy Governance &amp; Compliance</span></td>
                <td align="right" style="font-family:${F}; font-size:11px;"><a href="#unsubscribe" style="color:#94a3b8; text-decoration:underline;">Unsubscribe</a></td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>`;
  }

  /** Row counter for alternating backgrounds */
  private _rowIndex = 0;

  /** Builds a detail row for email tables (auto-escapes label and value) */
  private emailRow(label: string, value: string, valueStyle?: string): string {
    const bg = this._rowIndex % 2 === 0 ? '#f8fafc' : '#ffffff';
    this._rowIndex++;
    const F = "'Segoe UI', -apple-system, BlinkMacSystemFont, Roboto, Helvetica, Arial, sans-serif";
    return `<tr>
      <td width="38%" style="background-color:${bg}; padding:12px 20px; font-family:${F}; font-size:11px; font-weight:600; text-transform:uppercase; letter-spacing:0.8px; color:#64748b; border-bottom:1px solid #f1f5f9;">${escapeHtml(label)}</td>
      <td width="62%" style="background-color:${bg}; padding:12px 20px; font-family:${F}; font-size:13px; font-weight:500; color:#334155; border-bottom:1px solid #f1f5f9;${valueStyle || ''}">${escapeHtml(value)}</td>
    </tr>`;
  }

  /** Builds a detail table for email templates */
  private emailTable(rows: string): string {
    this._rowIndex = 0; // reset alternating rows
    return `<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="border:1px solid #e2e8f0; border-radius:8px; overflow:hidden;">${rows}</table>`;
  }

  /** Builds a policy highlight card (auto-escapes text content) */
  private policyCard(policyNumber: string, policyName: string, accent: string, detail: string): string {
    const F = "'Segoe UI', -apple-system, BlinkMacSystemFont, Roboto, Helvetica, Arial, sans-serif";
    return `<div style="background:#f8fafc;padding:16px 20px;border-radius:8px;border-left:4px solid ${accent};margin:20px 0;font-family:${F};">
      <p style="margin:0 0 4px;font-size:11px;color:#64748b;font-weight:600;letter-spacing:0.8px;text-transform:uppercase;">${escapeHtml(policyNumber)}</p>
      <p style="margin:0;font-size:15px;font-weight:600;color:#0f172a;">${escapeHtml(policyName)}</p>
      <p style="margin:8px 0 0;color:${accent};font-weight:600;font-size:13px;">${escapeHtml(detail)}</p>
    </div>`;
  }

  private buildNewPolicyEmail(policy: IPolicy): string {
    const policyUrl = `${this.siteUrl}/SitePages/PolicyDetails.aspx?policyId=${policy.Id}`;
    const rows = this.emailRow('Policy Number', policy.PolicyNumber || 'N/A')
      + this.emailRow('Policy Name', policy.PolicyName)
      + this.emailRow('Category', policy.PolicyCategory || 'General')
      + this.emailRow('Effective Date', policy.EffectiveDate ? new Date(policy.EffectiveDate).toLocaleDateString() : 'Immediate');
    const summary = policy.Description
      ? `<div style="background:#f0fdfa;padding:16px;border-radius:8px;margin:16px 0;border:1px solid #ccfbf1;">
           <p style="margin:0 0 6px;font-weight:600;color:#0f766e;font-size:13px;">Summary</p>
           <p style="margin:0;color:#334155;font-size:14px;line-height:1.6;">${escapeHtml(policy.Description)}</p>
         </div>` : '';
    return this.buildEmailShell({
      headerGradient: 'linear-gradient(135deg, #0d9488 0%, #0f766e 100%)',
      headerIcon: '\u{1F4CB}', headerTitle: 'New Policy Published',
      content: `<p style="margin:0 0 16px;color:#334155;font-size:15px;line-height:1.6;">A new policy has been published that requires your attention.</p>
        ${this.emailTable(rows)}${summary}`,
      footerText: 'This is an automated notification from the DWx Policy Management System.',
      ctaUrl: policyUrl, ctaLabel: 'View Policy', ctaColor: '#0d9488',
    });
  }

  private buildPolicyUpdateEmail(policy: IPolicy, changeDescription: string): string {
    const policyUrl = `${this.siteUrl}/SitePages/PolicyDetails.aspx?policyId=${policy.Id}`;
    const rows = this.emailRow('Policy', `${policy.PolicyNumber} &mdash; ${policy.PolicyName}`)
      + this.emailRow('New Version', `v${policy.VersionNumber || '1.0'}`, 'font-weight:600;color:#0d9488;');
    return this.buildEmailShell({
      headerGradient: 'linear-gradient(135deg, #d97706 0%, #b45309 100%)',
      headerIcon: '\u{1F504}', headerTitle: 'Policy Updated',
      content: `<p style="margin:0 0 16px;color:#334155;font-size:15px;line-height:1.6;">A policy you are required to acknowledge has been updated.</p>
        ${this.emailTable(rows)}
        <div style="background:#fef3c7;border-left:4px solid #d97706;padding:16px;border-radius:0 8px 8px 0;margin:20px 0;">
          <p style="margin:0 0 6px;font-weight:700;color:#92400e;font-size:13px;">What Changed</p>
          <p style="margin:0;color:#451a03;font-size:14px;line-height:1.5;">${escapeHtml(changeDescription || 'The policy has been revised. Please review the updated content.')}</p>
        </div>
        <p style="margin:16px 0 0;color:#dc2626;font-weight:600;font-size:14px;">\u26A0\uFE0F You may need to re-acknowledge this policy.</p>`,
      footerText: 'This is an automated notification from the DWx Policy Management System.',
      ctaUrl: policyUrl, ctaLabel: 'Review Updated Policy', ctaColor: '#d97706',
    });
  }

  private buildAcknowledgementRequiredEmail(policy: IPolicy, acknowledgement: IPolicyAcknowledgement): string {
    const policyUrl = `${this.siteUrl}/SitePages/PolicyDetails.aspx?policyId=${policy.Id}`;
    const dueDate = acknowledgement.DueDate ? new Date(acknowledgement.DueDate).toLocaleDateString() : 'As soon as possible';
    const rows = this.emailRow('Policy', `${policy.PolicyNumber} &mdash; ${policy.PolicyName}`)
      + this.emailRow('Category', policy.PolicyCategory || 'General')
      + this.emailRow('Due Date', dueDate, acknowledgement.DueDate ? 'color:#dc2626;font-weight:700;' : '');
    return this.buildEmailShell({
      headerGradient: 'linear-gradient(135deg, #0d9488 0%, #0f766e 100%)',
      headerIcon: '\u{1F4DD}', headerTitle: 'Action Required: Policy Acknowledgement',
      content: `<p style="margin:0 0 16px;color:#334155;font-size:15px;line-height:1.6;">You are required to read and acknowledge the following policy:</p>
        ${this.emailTable(rows)}`,
      footerText: 'This is an automated notification from the DWx Policy Management System.<br>You will receive reminders if the policy is not acknowledged by the due date.',
      ctaUrl: policyUrl, ctaLabel: 'Read & Acknowledge Policy', ctaColor: '#0d9488',
    });
  }

  private buildReminder3DayEmail(policy: IPolicy, acknowledgement: IPolicyAcknowledgement): string {
    const policyUrl = `${this.siteUrl}/SitePages/PolicyDetails.aspx?policyId=${policy.Id}`;
    const dueDate = acknowledgement.DueDate ? new Date(acknowledgement.DueDate).toLocaleDateString() : 'Soon';
    return this.buildEmailShell({
      headerGradient: 'linear-gradient(135deg, #d97706 0%, #b45309 100%)',
      headerIcon: '\u23F0', headerTitle: 'Reminder: 3 Days Remaining',
      content: `<p style="margin:0 0 16px;color:#334155;font-size:15px;line-height:1.6;">This is a friendly reminder that you have <strong style="color:#d97706;">3 days</strong> remaining to acknowledge the following policy:</p>
        ${this.policyCard(policy.PolicyNumber || '', policy.PolicyName, '#d97706', `Due: ${dueDate}`)}`,
      footerText: 'This is reminder 1 of 2. You will receive a final reminder 1 day before the due date.',
      ctaUrl: policyUrl, ctaLabel: 'Acknowledge Now', ctaColor: '#d97706',
    });
  }

  private buildReminder1DayEmail(policy: IPolicy, acknowledgement: IPolicyAcknowledgement): string {
    const policyUrl = `${this.siteUrl}/SitePages/PolicyDetails.aspx?policyId=${policy.Id}`;
    const dueDate = acknowledgement.DueDate ? new Date(acknowledgement.DueDate).toLocaleDateString() : 'Tomorrow';
    return this.buildEmailShell({
      headerGradient: 'linear-gradient(135deg, #dc2626 0%, #991b1b 100%)',
      headerIcon: '\u{1F6A8}', headerTitle: 'Final Reminder: Due Tomorrow!',
      content: `<p style="margin:0 0 16px;color:#dc2626;font-weight:600;font-size:16px;">This policy acknowledgement is due <strong>TOMORROW</strong>. Please take action today.</p>
        ${this.policyCard(policy.PolicyNumber || '', policy.PolicyName, '#dc2626', `Due: ${dueDate}`)}`,
      footerText: 'Failure to acknowledge by the due date may result in compliance escalation to your manager.',
      ctaUrl: policyUrl, ctaLabel: 'Acknowledge Immediately', ctaColor: '#dc2626',
    });
  }

  private buildOverdueEmail(policy: IPolicy, acknowledgement: IPolicyAcknowledgement, daysOverdue: number): string {
    const policyUrl = `${this.siteUrl}/SitePages/PolicyDetails.aspx?policyId=${policy.Id}`;
    const wasDue = acknowledgement.DueDate ? new Date(acknowledgement.DueDate).toLocaleDateString() : 'N/A';
    return this.buildEmailShell({
      headerGradient: 'linear-gradient(135deg, #991b1b 0%, #7f1d1d 100%)',
      headerIcon: '\u{1F534}', headerTitle: 'OVERDUE: Immediate Action Required',
      bodyBg: '#fef2f2',
      content: `<p style="margin:0 0 16px;color:#991b1b;font-weight:700;font-size:17px;">Your policy acknowledgement is <strong>${daysOverdue} day${daysOverdue > 1 ? 's' : ''} OVERDUE</strong>.</p>
        ${this.policyCard(policy.PolicyNumber || '', policy.PolicyName, '#991b1b', `Was Due: ${wasDue}`)}
        <div style="background:#fef3c7;padding:16px;border-radius:8px;margin:20px 0;border:1px solid #fde68a;">
          <p style="margin:0;font-weight:700;color:#92400e;font-size:13px;">\u26A0\uFE0F Compliance Notice</p>
          <p style="margin:8px 0 0;color:#451a03;font-size:14px;line-height:1.5;">This overdue acknowledgement has been flagged in the compliance system and your manager has been notified.</p>
        </div>`,
      footerText: 'If you have questions about this policy, please contact your manager or HR.',
      ctaUrl: policyUrl, ctaLabel: 'Acknowledge Now', ctaColor: '#991b1b',
    });
  }

  private buildAcknowledgementCompleteEmail(policy: IPolicy, acknowledgement: IPolicyAcknowledgement): string {
    const certificateUrl = `${this.siteUrl}/SitePages/PolicyCertificate.aspx?acknowledgementId=${acknowledgement.Id}`;
    const ackDate = acknowledgement.AcknowledgedDate ? new Date(acknowledgement.AcknowledgedDate).toLocaleDateString() : new Date().toLocaleDateString();
    const rows = this.emailRow('Receipt Number', String(acknowledgement.Id))
      + this.emailRow('Policy Version', `v${policy.VersionNumber || '1.0'}`)
      + this.emailRow('Acknowledged', ackDate, 'color:#059669;font-weight:600;');
    return this.buildEmailShell({
      headerGradient: 'linear-gradient(135deg, #059669 0%, #047857 100%)',
      headerIcon: '\u2705', headerTitle: 'Policy Acknowledged Successfully',
      bodyBg: '#f0fdf4',
      content: `<p style="margin:0 0 16px;color:#334155;font-size:15px;line-height:1.6;">Thank you for acknowledging the following policy:</p>
        ${this.policyCard(policy.PolicyNumber || '', policy.PolicyName, '#059669', `\u2713 Acknowledged: ${ackDate}`)}
        ${this.emailTable(rows)}`,
      footerText: 'Please retain this email as confirmation of your policy acknowledgement.',
      ctaUrl: certificateUrl, ctaLabel: 'View Certificate', ctaColor: '#059669',
    });
  }

  private buildPolicyExpiringEmail(policy: IPolicy, daysUntilExpiry: number): string {
    const policyUrl = `${this.siteUrl}/SitePages/PolicyAdmin.aspx?policyId=${policy.Id}`;
    const rows = this.emailRow('Policy', `${policy.PolicyNumber} &mdash; ${policy.PolicyName}`)
      + this.emailRow('Expiry Date', policy.ExpiryDate ? new Date(policy.ExpiryDate).toLocaleDateString() : 'N/A', 'color:#d97706;font-weight:700;')
      + this.emailRow('Owner', policy.PolicyOwner?.Title || 'Unassigned');
    return this.buildEmailShell({
      headerGradient: 'linear-gradient(135deg, #d97706 0%, #b45309 100%)',
      headerIcon: '\u{1F4C5}', headerTitle: 'Policy Expiring Soon',
      content: `<p style="margin:0 0 16px;color:#334155;font-size:15px;line-height:1.6;">The following policy will expire in <strong style="color:#d97706;">${daysUntilExpiry} days</strong>:</p>
        ${this.emailTable(rows)}
        <div style="background:#fef3c7;padding:16px;border-radius:8px;margin:20px 0;border:1px solid #fde68a;">
          <p style="margin:0 0 8px;font-weight:700;color:#92400e;font-size:13px;">Action Required</p>
          <ul style="margin:0;padding-left:20px;color:#451a03;font-size:14px;line-height:1.8;">
            <li>Review the policy for accuracy and relevance</li>
            <li>Update and re-publish if changes are needed</li>
            <li>Extend the expiry date if the policy is still valid</li>
            <li>Retire the policy if no longer applicable</li>
          </ul>
        </div>`,
      footerText: 'This is an automated alert from the DWx Policy Management System.',
      ctaUrl: policyUrl, ctaLabel: 'Manage Policy', ctaColor: '#d97706',
    });
  }

  private buildPolicyApprovalEmail(policy: IPolicy, approverName: string): string {
    const policyUrl = `${this.siteUrl}/SitePages/PolicyDetails.aspx?policyId=${policy.Id}&mode=review`;
    const rows = this.emailRow('Policy', `${policy.PolicyNumber} &mdash; ${policy.PolicyName}`)
      + this.emailRow('Approved By', approverName)
      + this.emailRow('Published Date', new Date().toLocaleDateString());
    return this.buildEmailShell({
      headerGradient: 'linear-gradient(135deg, #059669 0%, #047857 100%)',
      headerIcon: '\u2705', headerTitle: 'Policy Approved',
      bodyBg: '#f0fdf4',
      content: `<p style="margin:0 0 16px;color:#334155;font-size:15px;line-height:1.6;">Great news! Your policy has been <strong style="color:#059669;">approved</strong> and is now published.</p>
        ${this.emailTable(rows)}`,
      footerText: 'This is an automated notification from the DWx Policy Management System.',
      ctaUrl: policyUrl, ctaLabel: 'View Published Policy', ctaColor: '#059669',
    });
  }

  private buildPolicyRejectionEmail(policy: IPolicy, approverName: string, reason: string): string {
    const policyUrl = `${this.siteUrl}/SitePages/PolicyBuilder.aspx?policyId=${policy.Id}`;

    return `
      <html>
        <body style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; color: #323130; max-width: 600px; margin: 0 auto;">
          <div style="background: linear-gradient(135deg, #d13438 0%, #a80000 100%); padding: 24px; border-radius: 8px 8px 0 0;">
            <h2 style="color: white; margin: 0;">⚠️ Policy Requires Revision</h2>
          </div>

          <div style="padding: 24px; background: #fde7e9; border-radius: 0 0 8px 8px;">
            <p>Your policy submission requires revisions before it can be approved.</p>

            <table style="width: 100%; border-collapse: collapse; margin: 20px 0; background: white; border-radius: 8px;">
              <tr>
                <td style="padding: 12px 16px; background: #f3f2f1; font-weight: 600;">Policy:</td>
                <td style="padding: 12px 16px;">${escapeHtml(policy.PolicyNumber || '')} - ${escapeHtml(policy.PolicyName)}</td>
              </tr>
              <tr>
                <td style="padding: 12px 16px; background: #f3f2f1; font-weight: 600;">Reviewer:</td>
                <td style="padding: 12px 16px;">${escapeHtml(approverName)}</td>
              </tr>
            </table>

            <div style="background: white; padding: 16px; border-radius: 8px; border-left: 4px solid #d13438; margin: 20px 0;">
              <strong>Feedback:</strong>
              <p style="margin: 8px 0 0 0;">${escapeHtml(reason || 'Please contact the reviewer for specific feedback.')}</p>
            </div>

            <p style="text-align: center; margin: 24px 0;">
              <a href="${policyUrl}"
                 style="background: #0078d4; color: white; padding: 14px 32px;
                        text-decoration: none; border-radius: 6px; display: inline-block;
                        font-weight: 600;">
                Edit Policy
              </a>
            </p>

            <p style="color: #605e5c; font-size: 12px; margin-top: 30px; text-align: center;">
              Please make the requested changes and resubmit for approval.
            </p>
          </div>
        </body>
      </html>
    `;
  }

  // ============================================================================
  // HELPER METHODS
  // ============================================================================

  private async sendNotification(notification: IPolicyNotification): Promise<void> {
    try {
      // Get recipient email if not provided
      if (!notification.recipientEmail) {
        const user = await this.sp.web.siteUsers.getById(notification.recipientId)();
        notification.recipientEmail = user.Email;
        notification.recipientName = user.Title;
      }

      if (!notification.recipientEmail) {
        logger.warn('PolicyNotificationService', 'No email found for recipient', { recipientId: notification.recipientId });
        return;
      }

      // Log the notification (in production, this would send via SP utility or Graph API)
      logger.info('PolicyNotificationService', 'Sending notification', {
        type: notification.notificationType,
        to: notification.recipientEmail,
        subject: notification.subject
      });

      // Store notification in local list for audit trail
      await this.logNotification(notification);

      // Fire cross-app notification to DWx Hub (if available)
      if (notification.sendInApp && this.dwxNotifications) {
        await this.sendCrossAppNotification(notification);
      }

    } catch (error) {
      logger.error('PolicyNotificationService', 'Failed to send notification:', error);
      throw error;
    }
  }

  /**
   * Send cross-app notification to DWx Hub DWX_Notifications list
   */
  private async sendCrossAppNotification(notification: IPolicyNotification): Promise<void> {
    try {
      if (!this.dwxNotifications) {
        return;
      }

      const policyDetailsUrl = `${this.siteUrl}/SitePages/PolicyDetails.aspx?policyId=${notification.policyId}`;

      await this.dwxNotifications.createNotification({
        Title: notification.subject,
        MessageBody: notification.body.replace(/<[^>]*>/g, '').substring(0, 500),
        NotificationType: this.mapToDwxNotificationType(notification.notificationType),
        Priority: this.mapToDwxPriority(notification.notificationType),
        SourceApp: 'PolicyManager',
        SourceItemId: notification.policyId || 0,
        SourceItemTitle: notification.subject,
        SourceItemUrl: policyDetailsUrl,
        RecipientEmail: notification.recipientEmail || '',
        Category: this.mapToDwxCategory(notification.notificationType),
        ActionUrl: policyDetailsUrl,
      });
    } catch (error) {
      // Non-blocking — cross-app notification failure should not break local flow
      logger.warn('PolicyNotificationService', 'Failed to send cross-app notification to DWx Hub:', error);
    }
  }

  private mapToDwxNotificationType(type: PolicyNotificationType): DwxNotificationType {
    const map: Record<PolicyNotificationType, DwxNotificationType> = {
      'NewPolicy': DwxNotificationType.PolicyPublished,
      'PolicyUpdated': DwxNotificationType.PolicyPublished,
      'AcknowledgementRequired': DwxNotificationType.AcknowledgementDue,
      'ApprovalRequired': DwxNotificationType.ApprovalRequired,
      'Reminder3Day': DwxNotificationType.AcknowledgementDue,
      'Reminder1Day': DwxNotificationType.AcknowledgementDue,
      'Overdue': DwxNotificationType.ComplianceAlert,
      'AcknowledgementComplete': DwxNotificationType.ApprovalCompleted,
      'PolicyExpiring': DwxNotificationType.PolicyExpiring,
      'PolicyApproved': DwxNotificationType.ApprovalCompleted,
      'PolicyRejected': DwxNotificationType.ApprovalCompleted,
      'DelegationRequest': DwxNotificationType.ApprovalRequired,
    };
    return map[type] || DwxNotificationType.SystemAlert;
  }

  private mapToDwxPriority(type: PolicyNotificationType): DwxNotificationPriority {
    if (['Overdue', 'PolicyExpiring', 'Reminder1Day'].includes(type)) return 'High';
    if (['AcknowledgementRequired', 'ApprovalRequired', 'Reminder3Day', 'PolicyApproved', 'PolicyRejected', 'DelegationRequest'].includes(type)) return 'Medium';
    return 'Low';
  }

  private mapToDwxCategory(type: PolicyNotificationType): DwxNotificationCategory {
    if (['PolicyApproved', 'PolicyRejected', 'ApprovalRequired', 'DelegationRequest'].includes(type)) return 'Approval';
    if (['Overdue', 'PolicyExpiring', 'AcknowledgementRequired', 'Reminder3Day', 'Reminder1Day'].includes(type)) return 'Compliance';
    if (['AcknowledgementComplete'].includes(type)) return 'Task';
    return 'Info';
  }

  private async logNotification(notification: IPolicyNotification): Promise<void> {
    try {
      await this.sp.web.lists
        .getByTitle('PM_Notifications')
        .items.add({
          Title: notification.subject,
          RecipientId: notification.recipientId,
          Message: notification.subject,
          Type: 'Policy',
          RelatedItemId: notification.policyId,
          Priority: 'Normal',
          IsRead: false
        });
    } catch (error) {
      // Log error but don't throw - notification logging is not critical
      logger.warn('PolicyNotificationService', 'Failed to log notification:', error);
    }
  }

  private async sendManagerOverdueAlert(
    policy: IPolicy,
    acknowledgement: IPolicyAcknowledgement,
    daysOverdue: number
  ): Promise<void> {
    const managerId = (acknowledgement as any).ManagerId;
    if (!managerId) return;

    const user = await this.sp.web.siteUsers.getById(acknowledgement.AckUserId)();

    const notification: IPolicyNotification = {
      recipientId: managerId,
      notificationType: 'Overdue',
      subject: `Team Member Overdue: Policy Acknowledgement - ${user.Title}`,
      body: this.buildManagerOverdueEmail(policy, user.Title, daysOverdue),
      policyId: policy.Id,
      sendEmail: true,
      sendInApp: true
    };

    await this.sendNotification(notification);
  }

  private buildManagerOverdueEmail(policy: IPolicy, employeeName: string, daysOverdue: number): string {
    return `
      <html>
        <body style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; color: #323130; max-width: 600px; margin: 0 auto;">
          <div style="background: linear-gradient(135deg, #d13438 0%, #a80000 100%); padding: 24px; border-radius: 8px 8px 0 0;">
            <h2 style="color: white; margin: 0;">📊 Team Compliance Alert</h2>
          </div>

          <div style="padding: 24px; background: #f9f9f9; border-radius: 0 0 8px 8px;">
            <p>A member of your team has an overdue policy acknowledgement:</p>

            <table style="width: 100%; border-collapse: collapse; margin: 20px 0; background: white; border-radius: 8px;">
              <tr>
                <td style="padding: 12px 16px; background: #f3f2f1; font-weight: 600;">Employee:</td>
                <td style="padding: 12px 16px;">${escapeHtml(employeeName)}</td>
              </tr>
              <tr>
                <td style="padding: 12px 16px; background: #f3f2f1; font-weight: 600;">Policy:</td>
                <td style="padding: 12px 16px;">${escapeHtml(policy.PolicyNumber || '')} - ${escapeHtml(policy.PolicyName)}</td>
              </tr>
              <tr>
                <td style="padding: 12px 16px; background: #f3f2f1; font-weight: 600;">Days Overdue:</td>
                <td style="padding: 12px 16px; color: #d13438; font-weight: 600;">${daysOverdue}</td>
              </tr>
            </table>

            <p>Please follow up with your team member to ensure compliance.</p>

            <p style="color: #605e5c; font-size: 12px; margin-top: 30px; text-align: center;">
              This is an automated compliance alert from the JML Policy Management System.
            </p>
          </div>
        </body>
      </html>
    `;
  }


  private async updateReminderSchedule(scheduleId: number, updates: Partial<IReminderSchedule>): Promise<void> {
    try {
      await this.sp.web.lists
        .getByTitle(this.REMINDER_SCHEDULE_LIST)
        .items.getById(scheduleId)
        .update(updates as any);
    } catch (error) {
      logger.error('PolicyNotificationService', 'Failed to update reminder schedule:', error);
    }
  }
}
