// @ts-nocheck
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
import { NotificationLists, PolicyLists } from '../constants/SharePointListNames';
import { DwxNotificationService } from '@dwx/core';

/**
 * Policy notification types
 */
export type PolicyNotificationType =
  | 'NewPolicy'
  | 'PolicyUpdated'
  | 'AcknowledgementRequired'
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
  private readonly NOTIFICATION_LIST = NotificationLists.POLICY_NOTIFICATIONS;
  private readonly REMINDER_SCHEDULE_LIST = NotificationLists.REMINDER_SCHEDULE;

  constructor(sp: SPFI, siteUrl: string, dwxNotificationService?: DwxNotificationService) {
    this.sp = sp;
    this.siteUrl = siteUrl;
    this.dwxNotifications = dwxNotificationService || null;
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
    try {
      for (const recipientId of recipientIds) {
        const notification: IPolicyNotification = {
          recipientId,
          notificationType: 'NewPolicy',
          subject: `New Policy Published: ${policy.PolicyName}`,
          body: this.buildNewPolicyEmail(policy),
          policyId: policy.Id,
          sendEmail: true,
          sendInApp: true
        };

        await this.sendNotification(notification);
      }
    } catch (error) {
      logger.error('PolicyNotificationService', 'Failed to send new policy notification:', error);
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
    try {
      for (const recipientId of recipientIds) {
        const notification: IPolicyNotification = {
          recipientId,
          notificationType: 'PolicyUpdated',
          subject: `Policy Updated: ${policy.PolicyName}`,
          body: this.buildPolicyUpdateEmail(policy, changeDescription),
          policyId: policy.Id,
          sendEmail: true,
          sendInApp: true
        };

        await this.sendNotification(notification);
      }
    } catch (error) {
      logger.error('PolicyNotificationService', 'Failed to send policy update notification:', error);
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
        .items.filter(`PolicyId eq ${policyId} and UserId eq ${userId}`)
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
      const threeDaysFromNow = new Date(now);
      threeDaysFromNow.setDate(threeDaysFromNow.getDate() + 3);

      const oneDayFromNow = new Date(now);
      oneDayFromNow.setDate(oneDayFromNow.getDate() + 1);

      // Get all pending schedules
      const schedules = await this.sp.web.lists
        .getByTitle(this.REMINDER_SCHEDULE_LIST)
        .items.filter(`(Reminder3DaySent eq false or Reminder1DaySent eq false or OverdueSent eq false)`)
        .top(500)() as any[];

      for (const schedule of schedules) {
        stats.processed++;
        const dueDate = new Date(schedule.DueDate);
        const daysToDue = Math.ceil((dueDate.getTime() - now.getTime()) / (1000 * 60 * 60 * 24));

        // Get policy and acknowledgement info
        const policy = await this.getPolicy(schedule.PolicyId);
        if (!policy) continue;

        const acknowledgement = await this.getAcknowledgement(schedule.PolicyId, schedule.UserId);
        if (!acknowledgement || acknowledgement.AckStatus === AcknowledgementStatus.Acknowledged) {
          // Already acknowledged, remove schedule
          await this.sp.web.lists
            .getByTitle(this.REMINDER_SCHEDULE_LIST)
            .items.getById(schedule.Id)
            .delete();
          continue;
        }

        // 3-day reminder
        if (daysToDue <= 3 && daysToDue > 1 && !schedule.Reminder3DaySent) {
          await this.sendReminder3DayNotification(policy, acknowledgement);
          await this.updateReminderSchedule(schedule.Id, { reminder3DaySent: true });
          stats.reminders3Day++;
        }

        // 1-day reminder
        if (daysToDue === 1 && !schedule.Reminder1DaySent) {
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
      }

      logger.info('PolicyNotificationService', `Processed ${stats.processed} reminder schedules`, stats);
      return stats;
    } catch (error) {
      logger.error('PolicyNotificationService', 'Failed to process reminders:', error);
      return stats;
    }
  }

  // ============================================================================
  // EMAIL TEMPLATE BUILDERS
  // ============================================================================

  private buildNewPolicyEmail(policy: IPolicy): string {
    const policyUrl = `${this.siteUrl}/SitePages/PolicyDetails.aspx?policyId=${policy.Id}`;

    return `
      <html>
        <body style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; color: #323130; max-width: 600px; margin: 0 auto;">
          <div style="background: linear-gradient(135deg, #0078d4 0%, #004578 100%); padding: 24px; border-radius: 8px 8px 0 0;">
            <h2 style="color: white; margin: 0;">📋 New Policy Published</h2>
          </div>

          <div style="padding: 24px; background: #f9f9f9; border-radius: 0 0 8px 8px;">
            <p>A new policy has been published that requires your attention.</p>

            <table style="width: 100%; border-collapse: collapse; margin: 20px 0; background: white; border-radius: 8px; overflow: hidden;">
              <tr>
                <td style="padding: 12px 16px; background: #f3f2f1; font-weight: 600; width: 140px;">Policy Number:</td>
                <td style="padding: 12px 16px;">${policy.PolicyNumber || 'N/A'}</td>
              </tr>
              <tr>
                <td style="padding: 12px 16px; background: #f3f2f1; font-weight: 600;">Policy Name:</td>
                <td style="padding: 12px 16px;">${policy.PolicyName}</td>
              </tr>
              <tr>
                <td style="padding: 12px 16px; background: #f3f2f1; font-weight: 600;">Category:</td>
                <td style="padding: 12px 16px;">${policy.PolicyCategory || 'General'}</td>
              </tr>
              <tr>
                <td style="padding: 12px 16px; background: #f3f2f1; font-weight: 600;">Effective Date:</td>
                <td style="padding: 12px 16px;">${policy.EffectiveDate ? new Date(policy.EffectiveDate).toLocaleDateString() : 'Immediate'}</td>
              </tr>
            </table>

            ${policy.Description ? `
              <div style="background: white; padding: 16px; border-radius: 8px; margin-bottom: 20px;">
                <strong>Summary:</strong>
                <p style="margin: 8px 0 0 0; color: #605e5c;">${policy.Description}</p>
              </div>
            ` : ''}

            <p style="text-align: center; margin: 24px 0;">
              <a href="${policyUrl}"
                 style="background: #0078d4; color: white; padding: 14px 32px;
                        text-decoration: none; border-radius: 6px; display: inline-block;
                        font-weight: 600;">
                View Policy
              </a>
            </p>

            <p style="color: #605e5c; font-size: 12px; margin-top: 30px; text-align: center;">
              This is an automated notification from the JML Policy Management System.
            </p>
          </div>
        </body>
      </html>
    `;
  }

  private buildPolicyUpdateEmail(policy: IPolicy, changeDescription: string): string {
    const policyUrl = `${this.siteUrl}/SitePages/PolicyDetails.aspx?policyId=${policy.Id}`;

    return `
      <html>
        <body style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; color: #323130; max-width: 600px; margin: 0 auto;">
          <div style="background: linear-gradient(135deg, #f7630c 0%, #ca5010 100%); padding: 24px; border-radius: 8px 8px 0 0;">
            <h2 style="color: white; margin: 0;">🔄 Policy Updated</h2>
          </div>

          <div style="padding: 24px; background: #f9f9f9; border-radius: 0 0 8px 8px;">
            <p>A policy you are required to acknowledge has been updated.</p>

            <table style="width: 100%; border-collapse: collapse; margin: 20px 0; background: white; border-radius: 8px;">
              <tr>
                <td style="padding: 12px 16px; background: #f3f2f1; font-weight: 600; width: 140px;">Policy:</td>
                <td style="padding: 12px 16px;">${policy.PolicyNumber} - ${policy.PolicyName}</td>
              </tr>
              <tr>
                <td style="padding: 12px 16px; background: #f3f2f1; font-weight: 600;">New Version:</td>
                <td style="padding: 12px 16px;">v${policy.VersionNumber || '1.0'}</td>
              </tr>
            </table>

            <div style="background: #fff4ce; border-left: 4px solid #f7630c; padding: 16px; border-radius: 4px; margin: 20px 0;">
              <strong>What Changed:</strong>
              <p style="margin: 8px 0 0 0;">${changeDescription || 'The policy has been revised. Please review the updated content.'}</p>
            </div>

            <p style="color: #d13438; font-weight: 600;">⚠️ You may need to re-acknowledge this policy.</p>

            <p style="text-align: center; margin: 24px 0;">
              <a href="${policyUrl}"
                 style="background: #f7630c; color: white; padding: 14px 32px;
                        text-decoration: none; border-radius: 6px; display: inline-block;
                        font-weight: 600;">
                Review Updated Policy
              </a>
            </p>

            <p style="color: #605e5c; font-size: 12px; margin-top: 30px; text-align: center;">
              This is an automated notification from the JML Policy Management System.
            </p>
          </div>
        </body>
      </html>
    `;
  }

  private buildAcknowledgementRequiredEmail(policy: IPolicy, acknowledgement: IPolicyAcknowledgement): string {
    const policyUrl = `${this.siteUrl}/SitePages/PolicyDetails.aspx?policyId=${policy.Id}`;
    const dueDate = acknowledgement.DueDate ? new Date(acknowledgement.DueDate).toLocaleDateString() : 'As soon as possible';

    return `
      <html>
        <body style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; color: #323130; max-width: 600px; margin: 0 auto;">
          <div style="background: linear-gradient(135deg, #0078d4 0%, #004578 100%); padding: 24px; border-radius: 8px 8px 0 0;">
            <h2 style="color: white; margin: 0;">📝 Action Required: Policy Acknowledgement</h2>
          </div>

          <div style="padding: 24px; background: #f9f9f9; border-radius: 0 0 8px 8px;">
            <p>You are required to read and acknowledge the following policy:</p>

            <table style="width: 100%; border-collapse: collapse; margin: 20px 0; background: white; border-radius: 8px;">
              <tr>
                <td style="padding: 12px 16px; background: #f3f2f1; font-weight: 600; width: 140px;">Policy:</td>
                <td style="padding: 12px 16px;">${policy.PolicyNumber} - ${policy.PolicyName}</td>
              </tr>
              <tr>
                <td style="padding: 12px 16px; background: #f3f2f1; font-weight: 600;">Category:</td>
                <td style="padding: 12px 16px;">${policy.PolicyCategory}</td>
              </tr>
              <tr>
                <td style="padding: 12px 16px; background: #f3f2f1; font-weight: 600;">Due Date:</td>
                <td style="padding: 12px 16px; color: ${acknowledgement.DueDate ? '#d13438' : 'inherit'}; font-weight: 600;">
                  ${dueDate}
                </td>
              </tr>
            </table>

            <p style="text-align: center; margin: 24px 0;">
              <a href="${policyUrl}"
                 style="background: #0078d4; color: white; padding: 14px 32px;
                        text-decoration: none; border-radius: 6px; display: inline-block;
                        font-weight: 600;">
                Read & Acknowledge Policy
              </a>
            </p>

            <p style="color: #605e5c; font-size: 12px; margin-top: 30px; text-align: center;">
              This is an automated notification from the JML Policy Management System.<br>
              You will receive reminders if the policy is not acknowledged by the due date.
            </p>
          </div>
        </body>
      </html>
    `;
  }

  private buildReminder3DayEmail(policy: IPolicy, acknowledgement: IPolicyAcknowledgement): string {
    const policyUrl = `${this.siteUrl}/SitePages/PolicyDetails.aspx?policyId=${policy.Id}`;
    const dueDate = acknowledgement.DueDate ? new Date(acknowledgement.DueDate).toLocaleDateString() : 'Soon';

    return `
      <html>
        <body style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; color: #323130; max-width: 600px; margin: 0 auto;">
          <div style="background: linear-gradient(135deg, #f7630c 0%, #ca5010 100%); padding: 24px; border-radius: 8px 8px 0 0;">
            <h2 style="color: white; margin: 0;">⏰ Reminder: 3 Days Remaining</h2>
          </div>

          <div style="padding: 24px; background: #f9f9f9; border-radius: 0 0 8px 8px;">
            <p>This is a friendly reminder that you have <strong>3 days</strong> remaining to acknowledge the following policy:</p>

            <div style="background: white; padding: 20px; border-radius: 8px; border-left: 4px solid #f7630c; margin: 20px 0;">
              <h3 style="margin: 0 0 8px 0; color: #0078d4;">${policy.PolicyNumber}</h3>
              <p style="margin: 0; font-size: 16px;">${policy.PolicyName}</p>
              <p style="margin: 12px 0 0 0; color: #f7630c; font-weight: 600;">Due: ${dueDate}</p>
            </div>

            <p style="text-align: center; margin: 24px 0;">
              <a href="${policyUrl}"
                 style="background: #f7630c; color: white; padding: 14px 32px;
                        text-decoration: none; border-radius: 6px; display: inline-block;
                        font-weight: 600;">
                Acknowledge Now
              </a>
            </p>

            <p style="color: #605e5c; font-size: 12px; margin-top: 30px; text-align: center;">
              This is reminder 1 of 2. You will receive a final reminder 1 day before the due date.
            </p>
          </div>
        </body>
      </html>
    `;
  }

  private buildReminder1DayEmail(policy: IPolicy, acknowledgement: IPolicyAcknowledgement): string {
    const policyUrl = `${this.siteUrl}/SitePages/PolicyDetails.aspx?policyId=${policy.Id}`;
    const dueDate = acknowledgement.DueDate ? new Date(acknowledgement.DueDate).toLocaleDateString() : 'Tomorrow';

    return `
      <html>
        <body style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; color: #323130; max-width: 600px; margin: 0 auto;">
          <div style="background: linear-gradient(135deg, #d13438 0%, #a80000 100%); padding: 24px; border-radius: 8px 8px 0 0;">
            <h2 style="color: white; margin: 0;">🚨 Final Reminder: Due Tomorrow!</h2>
          </div>

          <div style="padding: 24px; background: #f9f9f9; border-radius: 0 0 8px 8px;">
            <p style="color: #d13438; font-weight: 600; font-size: 16px;">
              This policy acknowledgement is due <strong>TOMORROW</strong>. Please take action today.
            </p>

            <div style="background: white; padding: 20px; border-radius: 8px; border-left: 4px solid #d13438; margin: 20px 0;">
              <h3 style="margin: 0 0 8px 0; color: #0078d4;">${policy.PolicyNumber}</h3>
              <p style="margin: 0; font-size: 16px;">${policy.PolicyName}</p>
              <p style="margin: 12px 0 0 0; color: #d13438; font-weight: 600;">Due: ${dueDate}</p>
            </div>

            <p style="text-align: center; margin: 24px 0;">
              <a href="${policyUrl}"
                 style="background: #d13438; color: white; padding: 14px 32px;
                        text-decoration: none; border-radius: 6px; display: inline-block;
                        font-weight: 600;">
                Acknowledge Immediately
              </a>
            </p>

            <p style="color: #605e5c; font-size: 12px; margin-top: 30px; text-align: center;">
              Failure to acknowledge by the due date may result in compliance escalation to your manager.
            </p>
          </div>
        </body>
      </html>
    `;
  }

  private buildOverdueEmail(policy: IPolicy, acknowledgement: IPolicyAcknowledgement, daysOverdue: number): string {
    const policyUrl = `${this.siteUrl}/SitePages/PolicyDetails.aspx?policyId=${policy.Id}`;

    return `
      <html>
        <body style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; color: #323130; max-width: 600px; margin: 0 auto;">
          <div style="background: linear-gradient(135deg, #a80000 0%, #6e0000 100%); padding: 24px; border-radius: 8px 8px 0 0;">
            <h2 style="color: white; margin: 0;">🔴 OVERDUE: Immediate Action Required</h2>
          </div>

          <div style="padding: 24px; background: #fde7e9; border-radius: 0 0 8px 8px;">
            <p style="color: #a80000; font-weight: 600; font-size: 16px;">
              Your policy acknowledgement is <strong>${daysOverdue} day${daysOverdue > 1 ? 's' : ''} OVERDUE</strong>.
            </p>

            <div style="background: white; padding: 20px; border-radius: 8px; border-left: 4px solid #a80000; margin: 20px 0;">
              <h3 style="margin: 0 0 8px 0; color: #a80000;">${policy.PolicyNumber}</h3>
              <p style="margin: 0; font-size: 16px;">${policy.PolicyName}</p>
              <p style="margin: 12px 0 0 0; color: #a80000; font-weight: 600;">
                Was Due: ${acknowledgement.DueDate ? new Date(acknowledgement.DueDate).toLocaleDateString() : 'N/A'}
              </p>
            </div>

            <div style="background: #fff4ce; padding: 16px; border-radius: 8px; margin: 20px 0;">
              <strong>⚠️ Compliance Notice:</strong>
              <p style="margin: 8px 0 0 0;">This overdue acknowledgement has been flagged in the compliance system and your manager has been notified.</p>
            </div>

            <p style="text-align: center; margin: 24px 0;">
              <a href="${policyUrl}"
                 style="background: #a80000; color: white; padding: 14px 32px;
                        text-decoration: none; border-radius: 6px; display: inline-block;
                        font-weight: 600;">
                Acknowledge Now
              </a>
            </p>

            <p style="color: #605e5c; font-size: 12px; margin-top: 30px; text-align: center;">
              If you have questions about this policy, please contact your manager or HR.
            </p>
          </div>
        </body>
      </html>
    `;
  }

  private buildAcknowledgementCompleteEmail(policy: IPolicy, acknowledgement: IPolicyAcknowledgement): string {
    const certificateUrl = `${this.siteUrl}/SitePages/PolicyCertificate.aspx?acknowledgementId=${acknowledgement.Id}`;

    return `
      <html>
        <body style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; color: #323130; max-width: 600px; margin: 0 auto;">
          <div style="background: linear-gradient(135deg, #107c10 0%, #0b6a0b 100%); padding: 24px; border-radius: 8px 8px 0 0;">
            <h2 style="color: white; margin: 0;">✅ Policy Acknowledged Successfully</h2>
          </div>

          <div style="padding: 24px; background: #f9f9f9; border-radius: 0 0 8px 8px;">
            <p>Thank you for acknowledging the following policy:</p>

            <div style="background: white; padding: 20px; border-radius: 8px; border-left: 4px solid #107c10; margin: 20px 0;">
              <h3 style="margin: 0 0 8px 0; color: #0078d4;">${policy.PolicyNumber}</h3>
              <p style="margin: 0; font-size: 16px;">${policy.PolicyName}</p>
              <p style="margin: 12px 0 0 0; color: #107c10; font-weight: 600;">
                ✓ Acknowledged: ${acknowledgement.AcknowledgedDate ? new Date(acknowledgement.AcknowledgedDate).toLocaleDateString() : new Date().toLocaleDateString()}
              </p>
            </div>

            <table style="width: 100%; border-collapse: collapse; margin: 20px 0; background: white; border-radius: 8px;">
              <tr>
                <td style="padding: 12px 16px; background: #f3f2f1; font-weight: 600;">Receipt Number:</td>
                <td style="padding: 12px 16px;">${acknowledgement.Id}</td>
              </tr>
              <tr>
                <td style="padding: 12px 16px; background: #f3f2f1; font-weight: 600;">Policy Version:</td>
                <td style="padding: 12px 16px;">v${policy.VersionNumber || '1.0'}</td>
              </tr>
            </table>

            <p style="text-align: center; margin: 24px 0;">
              <a href="${certificateUrl}"
                 style="background: #107c10; color: white; padding: 14px 32px;
                        text-decoration: none; border-radius: 6px; display: inline-block;
                        font-weight: 600;">
                View Certificate
              </a>
            </p>

            <p style="color: #605e5c; font-size: 12px; margin-top: 30px; text-align: center;">
              Please retain this email as confirmation of your policy acknowledgement.
            </p>
          </div>
        </body>
      </html>
    `;
  }

  private buildPolicyExpiringEmail(policy: IPolicy, daysUntilExpiry: number): string {
    const policyUrl = `${this.siteUrl}/SitePages/PolicyAdmin.aspx?policyId=${policy.Id}`;

    return `
      <html>
        <body style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; color: #323130; max-width: 600px; margin: 0 auto;">
          <div style="background: linear-gradient(135deg, #f7630c 0%, #ca5010 100%); padding: 24px; border-radius: 8px 8px 0 0;">
            <h2 style="color: white; margin: 0;">📅 Policy Expiring Soon</h2>
          </div>

          <div style="padding: 24px; background: #f9f9f9; border-radius: 0 0 8px 8px;">
            <p>The following policy will expire in <strong>${daysUntilExpiry} days</strong>:</p>

            <table style="width: 100%; border-collapse: collapse; margin: 20px 0; background: white; border-radius: 8px;">
              <tr>
                <td style="padding: 12px 16px; background: #f3f2f1; font-weight: 600;">Policy:</td>
                <td style="padding: 12px 16px;">${policy.PolicyNumber} - ${policy.PolicyName}</td>
              </tr>
              <tr>
                <td style="padding: 12px 16px; background: #f3f2f1; font-weight: 600;">Expiry Date:</td>
                <td style="padding: 12px 16px; color: #f7630c; font-weight: 600;">
                  ${policy.ExpiryDate ? new Date(policy.ExpiryDate).toLocaleDateString() : 'N/A'}
                </td>
              </tr>
              <tr>
                <td style="padding: 12px 16px; background: #f3f2f1; font-weight: 600;">Owner:</td>
                <td style="padding: 12px 16px;">${policy.PolicyOwner?.Title || 'Unassigned'}</td>
              </tr>
            </table>

            <div style="background: #fff4ce; padding: 16px; border-radius: 8px; margin: 20px 0;">
              <strong>Action Required:</strong>
              <ul style="margin: 8px 0 0 0; padding-left: 20px;">
                <li>Review the policy for accuracy and relevance</li>
                <li>Update and re-publish if changes are needed</li>
                <li>Extend the expiry date if the policy is still valid</li>
                <li>Retire the policy if no longer applicable</li>
              </ul>
            </div>

            <p style="text-align: center; margin: 24px 0;">
              <a href="${policyUrl}"
                 style="background: #f7630c; color: white; padding: 14px 32px;
                        text-decoration: none; border-radius: 6px; display: inline-block;
                        font-weight: 600;">
                Manage Policy
              </a>
            </p>

            <p style="color: #605e5c; font-size: 12px; margin-top: 30px; text-align: center;">
              This is an automated alert from the JML Policy Management System.
            </p>
          </div>
        </body>
      </html>
    `;
  }

  private buildPolicyApprovalEmail(policy: IPolicy, approverName: string): string {
    const policyUrl = `${this.siteUrl}/SitePages/PolicyDetails.aspx?policyId=${policy.Id}`;

    return `
      <html>
        <body style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; color: #323130; max-width: 600px; margin: 0 auto;">
          <div style="background: linear-gradient(135deg, #107c10 0%, #0b6a0b 100%); padding: 24px; border-radius: 8px 8px 0 0;">
            <h2 style="color: white; margin: 0;">✅ Policy Approved</h2>
          </div>

          <div style="padding: 24px; background: #dff6dd; border-radius: 0 0 8px 8px;">
            <p>Great news! Your policy has been <strong style="color: #107c10;">approved</strong> and is now published.</p>

            <table style="width: 100%; border-collapse: collapse; margin: 20px 0; background: white; border-radius: 8px;">
              <tr>
                <td style="padding: 12px 16px; background: #f3f2f1; font-weight: 600;">Policy:</td>
                <td style="padding: 12px 16px;">${policy.PolicyNumber} - ${policy.PolicyName}</td>
              </tr>
              <tr>
                <td style="padding: 12px 16px; background: #f3f2f1; font-weight: 600;">Approved By:</td>
                <td style="padding: 12px 16px;">${approverName}</td>
              </tr>
              <tr>
                <td style="padding: 12px 16px; background: #f3f2f1; font-weight: 600;">Published Date:</td>
                <td style="padding: 12px 16px;">${new Date().toLocaleDateString()}</td>
              </tr>
            </table>

            <p style="text-align: center; margin: 24px 0;">
              <a href="${policyUrl}"
                 style="background: #107c10; color: white; padding: 14px 32px;
                        text-decoration: none; border-radius: 6px; display: inline-block;
                        font-weight: 600;">
                View Published Policy
              </a>
            </p>
          </div>
        </body>
      </html>
    `;
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
                <td style="padding: 12px 16px;">${policy.PolicyNumber} - ${policy.PolicyName}</td>
              </tr>
              <tr>
                <td style="padding: 12px 16px; background: #f3f2f1; font-weight: 600;">Reviewer:</td>
                <td style="padding: 12px 16px;">${approverName}</td>
              </tr>
            </table>

            <div style="background: white; padding: 16px; border-radius: 8px; border-left: 4px solid #d13438; margin: 20px 0;">
              <strong>Feedback:</strong>
              <p style="margin: 8px 0 0 0;">${reason || 'Please contact the reviewer for specific feedback.'}</p>
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
        RecipientEmail: notification.recipientEmail,
        Category: this.mapToDwxCategory(notification.notificationType),
        ActionUrl: policyDetailsUrl,
      });
    } catch (error) {
      // Non-blocking — cross-app notification failure should not break local flow
      logger.warn('PolicyNotificationService', 'Failed to send cross-app notification to DWx Hub:', error);
    }
  }

  private mapToDwxNotificationType(type: PolicyNotificationType): string {
    const map = {
      'NewPolicy': 'PolicyPublished',
      'PolicyUpdated': 'PolicyPublished',
      'AcknowledgementRequired': 'AcknowledgementDue',
      'Reminder3Day': 'AcknowledgementDue',
      'Reminder1Day': 'AcknowledgementDue',
      'Overdue': 'ComplianceAlert',
      'AcknowledgementComplete': 'ApprovalCompleted',
      'PolicyExpiring': 'PolicyExpiring',
      'PolicyApproved': 'ApprovalCompleted',
      'PolicyRejected': 'ApprovalCompleted',
      'DelegationRequest': 'ApprovalRequired',
    };
    return map[type] || 'SystemAlert';
  }

  private mapToDwxPriority(type: PolicyNotificationType): string {
    if (['Overdue', 'PolicyExpiring', 'Reminder1Day'].includes(type)) return 'High';
    if (['AcknowledgementRequired', 'Reminder3Day', 'PolicyApproved', 'PolicyRejected', 'DelegationRequest'].includes(type)) return 'Medium';
    return 'Low';
  }

  private mapToDwxCategory(type: PolicyNotificationType): string {
    if (['PolicyApproved', 'PolicyRejected', 'DelegationRequest'].includes(type)) return 'Approval';
    if (['Overdue', 'PolicyExpiring', 'AcknowledgementRequired', 'Reminder3Day', 'Reminder1Day'].includes(type)) return 'Compliance';
    if (['AcknowledgementComplete'].includes(type)) return 'Task';
    return 'Info';
  }

  private async logNotification(notification: IPolicyNotification): Promise<void> {
    try {
      await this.sp.web.lists
        .getByTitle(this.NOTIFICATION_LIST)
        .items.add({
          Title: notification.subject,
          RecipientId: notification.recipientId,
          RecipientEmail: notification.recipientEmail,
          NotificationType: notification.notificationType,
          PolicyId: notification.policyId,
          SentDate: new Date().toISOString(),
          EmailSent: notification.sendEmail,
          InAppSent: notification.sendInApp
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
                <td style="padding: 12px 16px;">${employeeName}</td>
              </tr>
              <tr>
                <td style="padding: 12px 16px; background: #f3f2f1; font-weight: 600;">Policy:</td>
                <td style="padding: 12px 16px;">${policy.PolicyNumber} - ${policy.PolicyName}</td>
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

  private async getPolicy(policyId: number): Promise<IPolicy | null> {
    try {
      const policy = await this.sp.web.lists
        .getByTitle(PolicyLists.POLICIES)
        .items.getById(policyId)() as IPolicy;
      return policy;
    } catch {
      return null;
    }
  }

  private async getAcknowledgement(policyId: number, userId: number): Promise<IPolicyAcknowledgement | null> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(PolicyLists.POLICY_ACKNOWLEDGEMENTS)
        .items.filter(`PolicyId eq ${policyId} and AckUserId eq ${userId}`)
        .top(1)() as IPolicyAcknowledgement[];
      return items.length > 0 ? items[0] : null;
    } catch {
      return null;
    }
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
