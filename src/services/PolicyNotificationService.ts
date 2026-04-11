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
import { EmailTemplateBuilder } from '../utils/EmailTemplateBuilder';
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

      // Wrap body in email shell using centralized EmailTemplateBuilder
      // mode=review for reviewer emails, no mode for ack emails (default = ack flow)
      const isReviewEvent = ['review-required', 'approval-request', 'review-withdrawn'].includes(opts.eventName);
      const modeParam = isReviewEvent ? '&mode=review' : '';
      const policyUrl = opts.mergeData.PolicyUrl || `${this.siteUrl}/SitePages/PolicyDetails.aspx?policyId=${opts.policyId || 0}${modeParam}`;
      // Map eventName to EmailNotificationType; default to 'policy-published' if no match
      const typeMap: Record<string, string> = {
        'review-required': 'review-required', 'approval-request': 'approval-request',
        'approval-approved': 'approval-approved', 'approval-rejected': 'approval-rejected',
        'review-withdrawn': 'policy-updated', 'policy-published': 'policy-published',
        'ack-required': 'ack-required', 'reminder-3day': 'reminder-3day',
        'reminder-1day': 'reminder-1day', 'overdue': 'overdue',
        'ack-complete': 'ack-complete', 'policy-expiring': 'policy-expiring',
        'policy-retired': 'policy-retired', 'sla-breach': 'sla-breach', 'welcome': 'welcome'
      };
      // Strip any leading greeting from body — EmailTemplateBuilder adds its own "Hi {name},"
      body = body.replace(/^(\s*<p>\s*)?Hi\s+[^,<]+,?\s*(<\/p>\s*)?/i, '').trim();

      const emailType = (typeMap[opts.eventName] || 'policy-published') as import('../utils/EmailTemplateBuilder').EmailNotificationType;
      const htmlBody = EmailTemplateBuilder.build(emailType, {
        recipientName: opts.recipientName || 'Team Member',
        headerTitle: subject,
        bodyText: body,
        rows: [
          { label: 'Policy', value: escapeHtml(opts.mergeData.PolicyTitle || '') },
          ...(opts.mergeData.PolicyNumber ? [{ label: 'Policy Number', value: escapeHtml(opts.mergeData.PolicyNumber) }] : []),
          ...(opts.mergeData.Category ? [{ label: 'Category', value: escapeHtml(opts.mergeData.Category) }] : [])
        ],
        ctaText: 'View in Policy Manager',
        ctaUrl: policyUrl
      });

      // Queue to PM_NotificationQueue
      console.log(`[PolicyNotificationService] Writing to PM_NotificationQueue: ${opts.eventName} → ${opts.recipientEmail}`);
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
      const qId = qResult?.data?.Id || (qResult as any)?.Id;
      console.log(`[PolicyNotificationService] ✓ Queued to PM_NotificationQueue (ID: ${qId}) for ${opts.recipientEmail}`);
      try { if (qId) await this.sp.web.lists.getByTitle('PM_NotificationQueue').items.getById(qId).update({ QueueStatus: 'Pending' }); } catch { /* */ }
    } catch (err) {
      console.error(`[PolicyNotificationService] ✗ FAILED to queue email for ${opts.eventName}:`, err);
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
    // Deduplicate recipient IDs
    const uniqueIds: number[] = [];
    for (const id of recipientIds) { if (uniqueIds.indexOf(id) === -1) uniqueIds.push(id); }
    this.resetBatchTracking();
    console.log(`[PolicyNotificationService] sendNewPolicyNotification: ${uniqueIds.length} unique recipients (from ${recipientIds.length})`);
    let failCount = 0;
    for (const recipientId of uniqueIds) {
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
    const policyUrl = `${this.siteUrl}/SitePages/PolicyDetails.aspx?policyId=${policy.Id}&mode=review`;
    // Deduplicate reviewer IDs — prevents sending the same person multiple emails
    const uniqueIds: number[] = [];
    for (const id of reviewerIds) { if (uniqueIds.indexOf(id) === -1) uniqueIds.push(id); }
    console.log('[PolicyNotificationService] sendSubmittedForReviewNotification:', {
      policyId: policy.Id, policyName: policy.PolicyName,
      originalIds: reviewerIds.length, uniqueIds: uniqueIds.length,
      submitterName, siteUrl: this.siteUrl
    });

    // Send email + in-app notification to each unique reviewer
    let failCount = 0;
    const sentEmails = new Set<string>(); // Track emails already sent to prevent duplicates
    for (const reviewerId of uniqueIds) {
      try {
        // Resolve reviewer email/name
        const reviewer = await this.sp.web.siteUsers.getById(reviewerId).select('Email', 'Title')();
        console.log(`[PolicyNotificationService] Reviewer ${reviewerId}:`, reviewer?.Email, reviewer?.Title);
        if (!reviewer?.Email) { console.warn(`[PolicyNotificationService] Reviewer ${reviewerId} has no email — skipping`); continue; }
        if (sentEmails.has(reviewer.Email.toLowerCase())) { console.log(`[PolicyNotificationService] Already sent to ${reviewer.Email} — skipping duplicate`); continue; }
        sentEmails.add(reviewer.Email.toLowerCase());

        // Queue templated email via PM_NotificationQueue
        console.log(`[PolicyNotificationService] Queuing review email for ${reviewer.Email}...`);
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
        console.log(`[PolicyNotificationService] ✓ Email queued for ${reviewer.Email}`);

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
  // EMAIL TEMPLATE BUILDERS — Now using centralized EmailTemplateBuilder
  // Private helpers (buildEmailShell, emailRow, emailTable, policyCard) removed;
  // all email rendering delegated to src/utils/EmailTemplateBuilder.ts
  // ============================================================================

  private buildNewPolicyEmail(policy: IPolicy): string {
    const policyUrl = `${this.siteUrl}/SitePages/PolicyDetails.aspx?policyId=${policy.Id}`;
    return EmailTemplateBuilder.policyPublished({
      recipientName: 'Team',
      policyTitle: policy.PolicyName,
      policyNumber: policy.PolicyNumber || 'N/A',
      publishedBy: policy.PolicyOwner?.Title || 'Policy Manager',
      category: policy.PolicyCategory || 'General',
      department: (policy as any).Department || '',
      riskLevel: policy.ComplianceRisk || 'Medium',
      effectiveDate: policy.EffectiveDate ? new Date(policy.EffectiveDate).toLocaleDateString() : 'Immediate',
      ctaUrl: policyUrl
    });
  }

  private buildPolicyUpdateEmail(policy: IPolicy, changeDescription: string): string {
    const policyUrl = `${this.siteUrl}/SitePages/PolicyDetails.aspx?policyId=${policy.Id}`;
    return EmailTemplateBuilder.policyUpdated({
      recipientName: 'Team',
      policyTitle: policy.PolicyName,
      policyNumber: policy.PolicyNumber || 'N/A',
      updatedBy: policy.PolicyOwner?.Title || 'Policy Manager',
      previousVersion: `v${String(Math.max(1, parseFloat(policy.VersionNumber || '1') - 1)).replace(/\.0+$/, '')}.0`,
      newVersion: `v${policy.VersionNumber || '1.0'}`,
      keyChanges: changeDescription || 'The policy has been revised. Please review the updated content.',
      ctaUrl: policyUrl
    });
  }

  private buildAcknowledgementRequiredEmail(policy: IPolicy, acknowledgement: IPolicyAcknowledgement): string {
    const policyUrl = `${this.siteUrl}/SitePages/PolicyDetails.aspx?policyId=${policy.Id}`;
    const dueDate = acknowledgement.DueDate ? new Date(acknowledgement.DueDate).toLocaleDateString() : 'As soon as possible';
    return EmailTemplateBuilder.ackRequired({
      recipientName: 'Team Member',
      policyTitle: policy.PolicyName,
      policyNumber: policy.PolicyNumber || 'N/A',
      assignedBy: policy.PolicyOwner?.Title || 'Policy Manager',
      category: policy.PolicyCategory || 'General',
      department: (policy as any).Department || '',
      riskLevel: policy.ComplianceRisk || 'Medium',
      dueDate,
      quizRequired: !!(policy as any).RequiresQuiz,
      ctaUrl: policyUrl
    });
  }

  private buildReminder3DayEmail(policy: IPolicy, acknowledgement: IPolicyAcknowledgement): string {
    const policyUrl = `${this.siteUrl}/SitePages/PolicyDetails.aspx?policyId=${policy.Id}`;
    const dueDate = acknowledgement.DueDate ? new Date(acknowledgement.DueDate).toLocaleDateString() : 'Soon';
    return EmailTemplateBuilder.reminder3Day({
      recipientName: 'Team Member',
      policyTitle: policy.PolicyName,
      policyNumber: policy.PolicyNumber || 'N/A',
      category: policy.PolicyCategory || 'General',
      dueDate,
      ctaUrl: policyUrl
    });
  }

  private buildReminder1DayEmail(policy: IPolicy, acknowledgement: IPolicyAcknowledgement): string {
    const policyUrl = `${this.siteUrl}/SitePages/PolicyDetails.aspx?policyId=${policy.Id}`;
    const dueDate = acknowledgement.DueDate ? new Date(acknowledgement.DueDate).toLocaleDateString() : 'Tomorrow';
    return EmailTemplateBuilder.reminder1Day({
      recipientName: 'Team Member',
      policyTitle: policy.PolicyName,
      policyNumber: policy.PolicyNumber || 'N/A',
      dueDate,
      ctaUrl: policyUrl
    });
  }

  private buildOverdueEmail(policy: IPolicy, acknowledgement: IPolicyAcknowledgement, daysOverdue: number): string {
    const policyUrl = `${this.siteUrl}/SitePages/PolicyDetails.aspx?policyId=${policy.Id}`;
    const wasDue = acknowledgement.DueDate ? new Date(acknowledgement.DueDate).toLocaleDateString() : 'N/A';
    return EmailTemplateBuilder.overdue({
      recipientName: 'Team Member',
      policyTitle: policy.PolicyName,
      policyNumber: policy.PolicyNumber || 'N/A',
      dueDate: wasDue,
      daysOverdue,
      escalatedTo: 'Your Manager',
      ctaUrl: policyUrl
    });
  }

  private buildAcknowledgementCompleteEmail(policy: IPolicy, acknowledgement: IPolicyAcknowledgement): string {
    const certificateUrl = `${this.siteUrl}/SitePages/PolicyCertificate.aspx?acknowledgementId=${acknowledgement.Id}`;
    const ackDate = acknowledgement.AcknowledgedDate ? new Date(acknowledgement.AcknowledgedDate).toLocaleDateString() : new Date().toLocaleDateString();
    return EmailTemplateBuilder.ackComplete({
      recipientName: 'Team Member',
      policyTitle: policy.PolicyName,
      policyNumber: policy.PolicyNumber || 'N/A',
      acknowledgedDate: ackDate,
      category: policy.PolicyCategory || 'General',
      ctaUrl: certificateUrl
    });
  }

  private buildPolicyExpiringEmail(policy: IPolicy, daysUntilExpiry: number): string {
    const policyUrl = `${this.siteUrl}/SitePages/PolicyAdmin.aspx?policyId=${policy.Id}`;
    return EmailTemplateBuilder.policyExpiring({
      recipientName: policy.PolicyOwner?.Title || 'Administrator',
      policyTitle: policy.PolicyName,
      policyNumber: policy.PolicyNumber || 'N/A',
      category: policy.PolicyCategory || 'General',
      currentVersion: `v${policy.VersionNumber || '1.0'}`,
      expiryDate: policy.ExpiryDate ? new Date(policy.ExpiryDate).toLocaleDateString() : 'N/A',
      daysUntilExpiry,
      ctaUrl: policyUrl
    });
  }

  private buildPolicyApprovalEmail(policy: IPolicy, approverName: string): string {
    const policyUrl = `${this.siteUrl}/SitePages/PolicyDetails.aspx?policyId=${policy.Id}&mode=review`;
    return EmailTemplateBuilder.approvalApproved({
      recipientName: 'Author',
      policyTitle: policy.PolicyName,
      policyNumber: policy.PolicyNumber || 'N/A',
      approvedBy: approverName,
      decisionDate: new Date().toLocaleDateString(),
      comments: 'No comments',
      ctaUrl: policyUrl
    });
  }

  private buildPolicyRejectionEmail(policy: IPolicy, approverName: string, reason: string): string {
    const policyUrl = `${this.siteUrl}/SitePages/PolicyBuilder.aspx?policyId=${policy.Id}`;
    return EmailTemplateBuilder.approvalRejected({
      recipientName: 'Author',
      policyTitle: policy.PolicyName,
      policyNumber: policy.PolicyNumber || 'N/A',
      rejectedBy: approverName,
      decisionDate: new Date().toLocaleDateString(),
      reason: reason || 'Please contact the reviewer for specific feedback.',
      ctaUrl: policyUrl
    });
  }

  // ============================================================================
  // HELPER METHODS
  // ============================================================================

  // Track emails sent in current batch to prevent duplicates within a single publish/notify operation
  private _sentEmailsThisBatch = new Set<string>();

  /** Reset batch tracking — call before starting a new notification batch */
  public resetBatchTracking(): void { this._sentEmailsThisBatch.clear(); }

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

      // Dedup: skip if already sent to this email in current batch
      const dedupKey = `${notification.notificationType}:${notification.recipientEmail.toLowerCase()}`;
      if (this._sentEmailsThisBatch.has(dedupKey)) {
        console.log(`[PolicyNotificationService] Dedup: skipping ${dedupKey} (already sent this batch)`);
        return;
      }
      this._sentEmailsThisBatch.add(dedupKey);

      // Log the notification (in production, this would send via SP utility or Graph API)
      logger.info('PolicyNotificationService', 'Sending notification', {
        type: notification.notificationType,
        to: notification.recipientEmail,
        subject: notification.subject
      });

      // Store notification in local list for audit trail
      await this.logNotification(notification);

      // Queue email to PM_NotificationQueue so the Logic App can send it
      // Guard: RecipientEmail MUST contain '@' — display names crash the Logic App
      if (notification.sendEmail && notification.recipientEmail && notification.recipientEmail.includes('@')) {
        try {
          await this.sp.web.lists.getByTitle('PM_NotificationQueue').items.add({
            Title: notification.subject,
            RecipientEmail: notification.recipientEmail,
            RecipientName: notification.recipientName || '',
            PolicyId: notification.policyId || 0,
            PolicyTitle: notification.subject,
            NotificationType: notification.notificationType,
            Channel: 'Email',
            Message: notification.body,
            QueueStatus: 'Pending',
            Priority: 'Normal'
          });
        } catch (emailQueueErr) {
          // Email queue failure must not break in-app notifications
          logger.warn('PolicyNotificationService', 'Failed to queue email to PM_NotificationQueue:', emailQueueErr);
        }
      }

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
